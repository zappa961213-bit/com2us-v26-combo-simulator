import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

const teamAlias: Record<string, string> = {
  해태: 'KIA',
  기아: 'KIA',
  빙그레: '한화',
  OB: '두산',
  현대: '키움',
  삼미: '키움',
  청보: '키움',
  태평양: '키움',
  SK: 'SSG',
  쌍방울: 'SSG',
  MBC: 'LG',
};

const dreamTeams = ['두산', '삼성', '롯데', 'SSG', 'KT'];
const nanumTeams = ['한화', 'KIA', '키움', 'LG', 'NC'];
const allTeams = ['두산', '삼성', '한화', '롯데', 'KIA', '키움', 'SSG', 'LG', 'NC', 'KT'];
const comboSignatureRate = 0.09;

type GameStage = 'ready' | 'open' | 'back' | 'shuffling' | 'shuffled' | 'picked';

interface CardData {
  id: string;
  type: 'combo' | 'normal';
  team: string;
  player: string;
  year: string;
  position: string;
  orderKey: number;
}

function normalizeTeam(team: string) {
  return teamAlias[team] || team;
}

function getLogoTeam(team: string) {
  const normalized = String(team).trim();

  switch (normalized) {
    case '두산':
      return 'doosan';
    case 'OB':
      return 'ob';
    case '삼성':
      return 'samsung';
    case '한화':
      return 'hanwha';
    case '빙그레':
      return 'binggrae';
    case '롯데':
      return 'lotte';
    case 'KIA':
    case '기아':
      return 'kia';
    case '해태':
      return 'haitai';
    case '키움':
      return 'kiwoom';
    case '삼미':
      return 'sammi';
    case '청보':
      return 'cheongbo';
    case '태평양':
      return 'taepyeongyang';
    case '현대':
      return 'hyundai';
    case 'SSG':
      return 'ssg';
    case 'SK':
      return 'sk';
    case '쌍방울':
      return 'ssangbangwool';
    case 'LG':
      return 'lg';
    case 'MBC':
      return 'mbc';
    case 'NC':
      return 'nc';
    case 'KT':
      return 'kt';
    default:
      return 'kia';
  }
}

function rollComboCard() {
  return Math.random() < comboSignatureRate;
}

function shuffleArray<T>(array: T[]) {
  const copied = [...array];

  for (let i = copied.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [copied[i], copied[j]] = [copied[j], copied[i]];
  }

  return copied;
}

function formatYear(value: unknown) {
  if (value === 0 || value === '0' || value === '00') return '00';
  return String(value || '').padStart(2, '0');
}

export default function App() {
  const [cards, setCards] = useState<CardData[]>([]);
  const [comboExclusivePool, setComboExclusivePool] = useState<CardData[]>([]);
  const [normalPool, setNormalPool] = useState<CardData[]>([]);
  const [selectedFilter, setSelectedFilter] = useState('전체');
  const [isRolling, setIsRolling] = useState(false);
  const [dbLoaded, setDbLoaded] = useState(false);
  const [stage, setStage] = useState<GameStage>('ready');
  const [pickedCardId, setPickedCardId] = useState<string | null>(null);

  useEffect(() => {
    async function loadExcelDatabase() {
      try {
        const response = await fetch('/카드DB.xlsx');

        if (!response.ok) {
          throw new Error('카드DB.xlsx 파일을 찾을 수 없음');
        }

        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        const comboSheet = workbook.Sheets['조합전용시그'];
        const normalSheet = workbook.Sheets['일반시그'];

        if (!comboSheet || !normalSheet) {
          throw new Error('시트 이름 확인 필요');
        }

        const comboData = XLSX.utils.sheet_to_json<any>(comboSheet);
        const normalData = XLSX.utils.sheet_to_json<any>(normalSheet);

        const mappedCombo: CardData[] = comboData.map((card: any, index: number) => ({
          id: `C-${index}`,
          type: 'combo',
          team: String(card.팀 || ''),
          player: String(card.선수명 || ''),
          year: formatYear(card.시즌),
          position: String(card.포지션 || ''),
          orderKey: index,
        }));

        const mappedNormal: CardData[] = normalData.map((card: any, index: number) => ({
          id: `N-${index}`,
          type: 'normal',
          team: String(card.팀 || ''),
          player: String(card.선수명 || ''),
          year: formatYear(card.시즌),
          position: String(card.포지션 || ''),
          orderKey: index,
        }));

        setComboExclusivePool(mappedCombo);
        setNormalPool(mappedNormal);
        setDbLoaded(true);
      } catch (error) {
        console.error('엑셀 로드 실패', error);
      }
    }

    loadExcelDatabase();
  }, []);

  const filteredNormalPool = useMemo(() => {
    if (selectedFilter === '전체') return normalPool;

    return normalPool.filter((card) => {
      const normalized = normalizeTeam(card.team);
      if (selectedFilter === '드림') return dreamTeams.includes(normalized);
      if (selectedFilter === '나눔') return nanumTeams.includes(normalized);
      return normalized === selectedFilter;
    });
  }, [normalPool, selectedFilter]);

  const filteredComboPool = useMemo(() => {
    if (selectedFilter === '전체') return comboExclusivePool;

    return comboExclusivePool.filter((card) => {
      const normalized = normalizeTeam(card.team);
      if (selectedFilter === '드림') return dreamTeams.includes(normalized);
      if (selectedFilter === '나눔') return nanumTeams.includes(normalized);
      return normalized === selectedFilter;
    });
  }, [comboExclusivePool, selectedFilter]);

  function getRandomCard(pool: CardData[], usedIds: Set<string>) {
    const available = pool.filter((card) => !usedIds.has(card.id));
    if (available.length === 0) return null;
    return available[Math.floor(Math.random() * available.length)];
  }

  async function simulateSignatureCombo() {
    if (!dbLoaded || filteredNormalPool.length === 0) return;

    setIsRolling(true);
    setPickedCardId(null);
    setStage('ready');

    const generated: CardData[] = [];
    const usedIds = new Set<string>();

    for (let i = 0; i < 5; i++) {
      const useCombo = rollComboCard();
      const pool = useCombo && filteredComboPool.length > 0 ? filteredComboPool : filteredNormalPool;
      const picked = getRandomCard(pool, usedIds);

      if (picked) {
        usedIds.add(picked.id);
        generated.push({ ...picked, orderKey: i });
      }
    }

    setCards(generated);
    setStage('open');
    setIsRolling(false);
  }

  async function startShuffle() {
    if (stage !== 'open') return;

    setIsRolling(true);
    setStage('back');

    await new Promise((resolve) => setTimeout(resolve, 650));
    setStage('shuffling');

    for (let i = 0; i < 8; i++) {
      await new Promise((resolve) => setTimeout(resolve, 180));
      setCards((prev) => shuffleArray(prev).map((card, index) => ({ ...card, orderKey: index })));
    }

    await new Promise((resolve) => setTimeout(resolve, 250));
    setStage('shuffled');
    setIsRolling(false);
  }

  function pickCard(cardId: string) {
    if (stage !== 'shuffled') return;
    setPickedCardId(cardId);
    setStage('picked');
  }

  function cardFaceVisible(cardId: string) {
    return stage === 'open' || stage === 'picked' || pickedCardId === cardId;
  }

  return (
    <div className="min-h-screen overflow-hidden bg-black text-white relative">
      <style>{`
        @keyframes shuffleShake {
          0% { transform: translateX(0) translateY(0) rotate(0deg); }
          20% { transform: translateX(-18px) translateY(10px) rotate(-5deg); }
          40% { transform: translateX(18px) translateY(-8px) rotate(5deg); }
          60% { transform: translateX(-10px) translateY(-14px) rotate(3deg); }
          80% { transform: translateX(14px) translateY(10px) rotate(-3deg); }
          100% { transform: translateX(0) translateY(0) rotate(0deg); }
        }
        .shuffle-card { animation: shuffleShake 0.36s ease-in-out infinite; }
      `}</style>

      <div className="absolute inset-0 bg-[radial-gradient(circle_at_top,#54200f_0%,#120301_42%,#000_100%)]" />
      <div className="absolute left-0 right-0 top-72 h-44 bg-gradient-to-r from-orange-500/20 via-yellow-300/20 to-orange-500/20 blur-3xl" />

      <main className="relative z-10 flex flex-col items-center px-4 py-8 gap-7">
        <section className="text-center space-y-2">
          <h1 className="text-4xl md:text-5xl font-black tracking-wide text-pink-300 drop-shadow-[0_0_18px_rgba(255,105,180,0.6)]">
            컴프야V26 시그 조합 시뮬레이터
          </h1>
          <p className="text-zinc-300">조합전용 시그 9% / 일반 시그 91%</p>
          <p className="text-sm text-zinc-500">
            {dbLoaded
              ? `DB 로딩 완료 / 조합전용 ${comboExclusivePool.length}장 / 일반 ${normalPool.length}장`
              : '엑셀 DB 로딩 중...'}
          </p>
        </section>

        <section className="flex flex-col items-center gap-3">
          <div className="text-sm font-bold text-zinc-400">팀 일반 확정권</div>
          <div className="flex gap-3 flex-wrap justify-center">
            {['전체', '드림', '나눔'].map((filter) => (
              <button
                key={filter}
                onClick={() => setSelectedFilter(filter)}
                disabled={isRolling}
                className={`px-6 py-3 rounded-xl font-black transition-all ${
                  selectedFilter === filter
                    ? 'bg-pink-500 text-white shadow-[0_0_18px_rgba(236,72,153,0.7)]'
                    : 'bg-zinc-800 text-zinc-300 hover:bg-zinc-700'
                }`}
              >
                {filter}
              </button>
            ))}
          </div>
        </section>

        <section className="flex flex-col items-center gap-3 max-w-4xl">
          <div className="text-sm font-bold text-zinc-400">팀 고급 확정권</div>
          <div className="flex gap-2 flex-wrap justify-center">
            {allTeams.map((team) => (
              <button
                key={team}
                onClick={() => setSelectedFilter(team)}
                disabled={isRolling}
                className={`px-4 py-2 rounded-xl text-sm font-black transition-all ${
                  selectedFilter === team
                    ? 'bg-orange-400 text-black shadow-[0_0_18px_rgba(251,146,60,0.75)]'
                    : 'bg-zinc-900 text-zinc-300 border border-zinc-700 hover:bg-zinc-800'
                }`}
              >
                {team}
              </button>
            ))}
          </div>
        </section>

        <button
          onClick={simulateSignatureCombo}
          disabled={isRolling || !dbLoaded}
          className="px-10 py-4 rounded-2xl bg-gradient-to-r from-pink-500 to-purple-600 text-xl font-black shadow-[0_0_25px_rgba(217,70,239,0.6)] hover:scale-105 transition-transform disabled:opacity-50"
        >
          {isRolling ? '진행 중...' : '시그 조합 실행'}
        </button>

        {stage === 'open' && (
          <button
            onClick={startShuffle}
            className="px-8 py-3 rounded-2xl bg-gradient-to-r from-yellow-400 to-orange-500 text-black font-black shadow-[0_0_20px_rgba(255,180,0,0.7)] hover:scale-105 transition-transform"
          >
            셔플 시작
          </button>
        )}

        <p className="h-6 text-sm text-zinc-400">
          {stage === 'open' && '카드 5장이 공개되었습니다. 확인 후 셔플을 시작하세요.'}
          {stage === 'back' && '카드를 뒤집는 중...'}
          {stage === 'shuffling' && '카드를 섞는 중...'}
          {stage === 'shuffled' && '뒷면 카드 1장을 선택하세요'}
          {stage === 'picked' && '선택 결과 공개'}
        </p>

        <section className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-6 max-w-7xl min-h-96">
          {cards.map((card) => {
            const isVisible = cardFaceVisible(card.id);
            const isPicked = pickedCardId === card.id;

            return (
              <button
                key={card.id}
                onClick={() => pickCard(card.id)}
                disabled={stage !== 'shuffled'}
                className={`relative w-32 h-48 sm:w-40 sm:h-60 lg:w-52 lg:h-80
  rounded-3xl transition-all duration-500 [perspective:1000px]
  ${stage === 'shuffled' ? 'hover:scale-105 cursor-pointer' : 'cursor-default'}
  ${
    stage === 'picked'
      ? isPicked
        ? 'scale-125 z-20'
        : 'scale-90 opacity-40 blur-[1px]'
      : ''
  }
  ${stage === 'shuffling' ? 'shuffle-card' : ''}
`}

                style={{ order: card.orderKey }}
              >
                <div
                  className={`relative w-full h-full transition-transform duration-700 [transform-style:preserve-3d] ${
                    isVisible ? '[transform:rotateY(0deg)]' : '[transform:rotateY(180deg)]'
                  }`}
                >
                  <div
                    className={`absolute inset-0 rounded-3xl overflow-hidden border-4 [backface-visibility:hidden] ${
                      card.type === 'combo'
                        ? 'border-orange-400 shadow-[0_0_45px_rgba(255,140,0,0.95)]'
                        : 'border-pink-300 shadow-[0_0_22px_rgba(255,105,180,0.55)]'
                    }`}
                  >
                    <div className="absolute inset-0 bg-gradient-to-b from-pink-200 via-pink-400 to-pink-100" />
                    <div className="absolute inset-0 bg-[radial-gradient(circle_at_center,rgba(255,255,255,0.55),transparent_55%)]" />
                    <div className="absolute inset-x-0 bottom-0 h-24 bg-white/40" />

                    {card.type === 'combo' && (
                      <div className="absolute inset-0 animate-pulse bg-orange-400/10" />
                    )}

                    <div className="relative z-10 flex h-full flex-col p-4">
                      <div className="flex items-start justify-between">
                        <div className="text-4xl font-black text-pink-700 leading-none">
                          {card.position}
                        </div>
                        <div className="h-12 w-12 rounded-full bg-white/95 flex items-center justify-center shadow-lg p-1 overflow-hidden">
                          <img
                            src={`/logos/${getLogoTeam(card.team)}.png`}
                            alt={card.team}
                            className="h-full w-full object-contain"
                          />
                        </div>
                      </div>

                      <div className="flex-1 flex items-center justify-center">
                        <div className="text-5xl font-black text-white/30 italic select-none">
                          Signature
                        </div>
                      </div>

                      <div className="text-center pb-1">
                        {card.type === 'combo' && (
                          <div className="text-orange-600 font-black text-sm tracking-widest mb-1">
                            조합전용
                          </div>
                        )}
                        <div className="text-3xl font-black text-zinc-950 tracking-tight">
                          {card.player}
                        </div>
                        <div className="text-xl font-black text-zinc-700">'{card.year}</div>
                      </div>
                    </div>
                  </div>

                  <div className="absolute inset-0 rounded-3xl overflow-hidden border-4 border-fuchsia-300 bg-gradient-to-br from-fuchsia-700 via-pink-500 to-purple-800 shadow-[0_0_30px_rgba(217,70,239,0.7)] [backface-visibility:hidden] [transform:rotateY(180deg)]">
                    <div className="absolute inset-3 rounded-2xl border border-white/30" />
                    <div className="absolute inset-0 bg-[radial-gradient(circle_at_center,rgba(255,255,255,0.25),transparent_55%)]" />
                    <div className="relative z-10 h-full flex items-center justify-center">
                      <div className="text-4xl font-black text-white/90 italic tracking-wider">
                        SIGN
                      </div>
                    </div>
                  </div>
                </div>
              </button>
            );
          })}
        </section>
      </main>
    </div>
  );
}
