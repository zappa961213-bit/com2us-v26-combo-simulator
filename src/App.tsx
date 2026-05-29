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
const autoCounts = [5, 10, 30, 50, 100];

type GameStage = 'ready' | 'open' | 'back' | 'shuffling' | 'shuffled' | 'picked';
type ComboMode = 'signature' | 'impact';
type MainTab = 'simulation' | 'wish' | 'history';
type WishSubTab = 'manage' | 'register';
type TicketType = 'normal' | 'advanced';
type AutoType = 'normal' | 'special';
type SpecialTarget = 'specific' | 'wish' | 'comboOnly';

interface CardData {
  id: string;
  mode: ComboMode;
  type: 'combo' | 'normal';
  team: string;
  player: string;
  year?: string;
  concept?: string;
  position: string;
  orderKey: number;
}

interface SimStats {
  totalCombos: number;
  normalTickets: number;
  advancedTickets: number;
}

interface HistoryState {
  cardCounts: Record<string, number>;
  normalTeamCounts: Record<string, number>;
  comboOnlyCount: number;
  wishHitCount: number;
}

interface AutoResultItem {
  card: CardData;
  count: number;
}

function normalizeTeam(team: string) {
  return teamAlias[String(team).trim()] || String(team).trim();
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

function formatYear(value: unknown) {
  if (value === 0 || value === '0' || value === '00') return '00';
  return String(value || '').padStart(2, '0');
}

function shuffleArray<T>(array: T[]) {
  const copied = [...array];

  for (let i = copied.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [copied[i], copied[j]] = [copied[j], copied[i]];
  }

  return copied;
}

function getCardLabel(card: CardData) {
  const title = card.mode === 'impact' ? card.concept || '' : card.year ? `'${card.year}` : '';
  return `${card.team} ${card.player} ${title} ${card.position}`.replace(/\s+/g, ' ').trim();
}

function safeJsonParse<T>(value: string | null, fallback: T): T {
  if (!value) return fallback;
  try {
    return JSON.parse(value) as T;
  } catch {
    return fallback;
  }
}

function getTicketType(filter: string): TicketType {
  return ['전체', '드림', '나눔'].includes(filter) ? 'normal' : 'advanced';
}

function matchesFilter(card: CardData, filter: string) {
  if (filter === '전체') return true;

  const normalized = normalizeTeam(card.team);
  if (filter === '드림') return dreamTeams.includes(normalized);
  if (filter === '나눔') return nanumTeams.includes(normalized);

  return normalized === filter;
}

function getRandomCard(pool: CardData[], usedIds: Set<string>) {
  const available = pool.filter((card) => !usedIds.has(card.id));
  if (available.length === 0) return null;
  return available[Math.floor(Math.random() * available.length)];
}

function buildResultSummary(cards: CardData[]) {
  const counts = new Map<string, AutoResultItem>();

  cards.forEach((card) => {
    const current = counts.get(card.id);
    if (current) {
      counts.set(card.id, { card, count: current.count + 1 });
    } else {
      counts.set(card.id, { card, count: 1 });
    }
  });

  return Array.from(counts.values()).sort((a, b) => b.count - a.count);
}

export default function App() {
  const [cards, setCards] = useState<CardData[]>([]);

  const [signatureComboPool, setSignatureComboPool] = useState<CardData[]>([]);
  const [signatureNormalPool, setSignatureNormalPool] = useState<CardData[]>([]);
  const [impactComboPool, setImpactComboPool] = useState<CardData[]>([]);
  const [impactNormalPool, setImpactNormalPool] = useState<CardData[]>([]);

  const [mainTab, setMainTab] = useState<MainTab>('simulation');
  const [wishSubTab, setWishSubTab] = useState<WishSubTab>('manage');

  const [comboMode, setComboMode] = useState<ComboMode>('signature');
  const [selectedFilter, setSelectedFilter] = useState('전체');
  const [isRolling, setIsRolling] = useState(false);
  const [dbLoaded, setDbLoaded] = useState(false);
  const [stage, setStage] = useState<GameStage>('ready');
  const [pickedCardId, setPickedCardId] = useState<string | null>(null);

  const [wishIds, setWishIds] = useState<string[]>([]);
  const [wishSearch, setWishSearch] = useState('');
  const [wishModeFilter, setWishModeFilter] = useState<'all' | ComboMode>('all');
  const [wishTypeFilter, setWishTypeFilter] = useState<'all' | 'combo' | 'normal'>('all');
  const [wishTeamFilter, setWishTeamFilter] = useState('전체');

  const [stats, setStats] = useState<SimStats>({
    totalCombos: 0,
    normalTickets: 0,
    advancedTickets: 0,
  });

  const [history, setHistory] = useState<HistoryState>({
    cardCounts: {},
    normalTeamCounts: {},
    comboOnlyCount: 0,
    wishHitCount: 0,
  });

  const [autoOpen, setAutoOpen] = useState(false);
  const [autoType, setAutoType] = useState<AutoType>('normal');
  const [autoMode, setAutoMode] = useState<ComboMode>('signature');
  const [autoFilter, setAutoFilter] = useState('전체');
  const [autoCount, setAutoCount] = useState(10);
  const [specialTarget, setSpecialTarget] = useState<SpecialTarget>('wish');
  const [specialSearch, setSpecialSearch] = useState('');
  const [autoResult, setAutoResult] = useState('');
  const [autoResultItems, setAutoResultItems] = useState<AutoResultItem[]>([]);

  useEffect(() => {
    setWishIds(safeJsonParse<string[]>(localStorage.getItem('wishIds'), []));
    setStats(
      safeJsonParse<SimStats>(localStorage.getItem('simStats'), {
        totalCombos: 0,
        normalTickets: 0,
        advancedTickets: 0,
      })
    );
    setHistory(
      safeJsonParse<HistoryState>(localStorage.getItem('history'), {
        cardCounts: {},
        normalTeamCounts: {},
        comboOnlyCount: 0,
        wishHitCount: 0,
      })
    );
  }, []);

  useEffect(() => {
    localStorage.setItem('wishIds', JSON.stringify(wishIds));
  }, [wishIds]);

  useEffect(() => {
    localStorage.setItem('simStats', JSON.stringify(stats));
  }, [stats]);

  useEffect(() => {
    localStorage.setItem('history', JSON.stringify(history));
  }, [history]);

  useEffect(() => {
    async function loadExcelDatabase() {
      try {
        const response = await fetch('/카드DB.xlsx');

        if (!response.ok) {
          throw new Error('카드DB.xlsx 파일을 찾을 수 없음');
        }

        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        const signatureComboSheet = workbook.Sheets['조합전용시그'];
        const signatureNormalSheet = workbook.Sheets['일반시그'];
        const impactComboSheet = workbook.Sheets['조합전용임팩'];
        const impactNormalSheet = workbook.Sheets['일반임팩'];

        if (!signatureComboSheet || !signatureNormalSheet) {
          throw new Error('시그 시트 이름 확인 필요');
        }

        const signatureComboData = XLSX.utils.sheet_to_json<any>(signatureComboSheet);
        const signatureNormalData = XLSX.utils.sheet_to_json<any>(signatureNormalSheet);
        const impactComboData = impactComboSheet ? XLSX.utils.sheet_to_json<any>(impactComboSheet) : [];
        const impactNormalData = impactNormalSheet ? XLSX.utils.sheet_to_json<any>(impactNormalSheet) : [];

        setSignatureComboPool(
          signatureComboData.map((card: any, index: number) => ({
            id: `SIG-C-${index}`,
            mode: 'signature',
            type: 'combo',
            team: String(card.팀 || ''),
            player: String(card.선수명 || ''),
            year: formatYear(card.시즌),
            position: String(card.포지션 || ''),
            orderKey: index,
          }))
        );

        setSignatureNormalPool(
          signatureNormalData.map((card: any, index: number) => ({
            id: `SIG-N-${index}`,
            mode: 'signature',
            type: 'normal',
            team: String(card.팀 || ''),
            player: String(card.선수명 || ''),
            year: formatYear(card.시즌),
            position: String(card.포지션 || ''),
            orderKey: index,
          }))
        );

        setImpactComboPool(
          impactComboData.map((card: any, index: number) => ({
            id: `IMP-C-${index}`,
            mode: 'impact',
            type: 'combo',
            team: String(card.팀 || ''),
            player: String(card.선수명 || ''),
            concept: String(card.컨셉 || ''),
            position: String(card.포지션 || ''),
            orderKey: index,
          }))
        );

        setImpactNormalPool(
          impactNormalData.map((card: any, index: number) => ({
            id: `IMP-N-${index}`,
            mode: 'impact',
            type: 'normal',
            team: String(card.팀 || ''),
            player: String(card.선수명 || ''),
            concept: String(card.컨셉 || ''),
            position: String(card.포지션 || ''),
            orderKey: index,
          }))
        );

        setDbLoaded(true);
      } catch (error) {
        console.error('엑셀 로드 실패', error);
      }
    }

    loadExcelDatabase();
  }, []);

  const allPool = useMemo(
    () => [...signatureComboPool, ...signatureNormalPool, ...impactComboPool, ...impactNormalPool],
    [signatureComboPool, signatureNormalPool, impactComboPool, impactNormalPool]
  );

  const wishCards = useMemo(() => {
    return wishIds
      .map((id) => allPool.find((card) => card.id === id))
      .filter((card): card is CardData => Boolean(card));
  }, [wishIds, allPool]);

  const currentNormalPool = comboMode === 'signature' ? signatureNormalPool : impactNormalPool;

  const filteredNormalPool = useMemo(() => {
    return currentNormalPool.filter((card) => matchesFilter(card, selectedFilter));
  }, [currentNormalPool, selectedFilter]);

  const wishSearchResults = useMemo(() => {
    const keyword = wishSearch.trim().toLowerCase();

    return allPool
      .filter((card) => {
        if (wishModeFilter !== 'all' && card.mode !== wishModeFilter) return false;
        if (wishTypeFilter !== 'all' && card.type !== wishTypeFilter) return false;
        if (wishTeamFilter !== '전체' && normalizeTeam(card.team) !== wishTeamFilter) return false;

        if (!keyword) return true;

        const target = [
          card.team,
          normalizeTeam(card.team),
          card.player,
          card.year,
          card.concept,
          card.position,
          card.mode === 'signature' ? '시그니처 시그' : '임팩트 임팩',
          card.type === 'combo' ? '조합전용 조합 전용' : '일반',
        ]
          .join(' ')
          .toLowerCase();

        return target.includes(keyword);
      })
      .slice(0, 120);
  }, [allPool, wishSearch, wishModeFilter, wishTypeFilter, wishTeamFilter]);

  const historyCards = useMemo(() => {
    return Object.entries(history.cardCounts)
      .map(([id, count]) => {
        const card = allPool.find((item) => item.id === id);
        return card ? { card, count } : null;
      })
      .filter((item): item is { card: CardData; count: number } => Boolean(item))
      .sort((a, b) => b.count - a.count);
  }, [history.cardCounts, allPool]);

  function drawFive(mode: ComboMode, filter: string) {
    const normalPool = mode === 'signature' ? signatureNormalPool : impactNormalPool;
    const comboPool = mode === 'signature' ? signatureComboPool : impactComboPool;
    const rate = mode === 'signature' ? 0.09 : 0.15;

    const filteredNormal = normalPool.filter((card) => matchesFilter(card, filter));
    const filteredCombo = comboPool.filter((card) => matchesFilter(card, filter));

    if (filteredNormal.length === 0) return [];

    const generated: CardData[] = [];
    const usedIds = new Set<string>();

    for (let i = 0; i < 5; i++) {
      const useCombo = Math.random() < rate;
      const pool = useCombo && filteredCombo.length > 0 ? filteredCombo : filteredNormal;
      const picked = getRandomCard(pool, usedIds);

      if (picked) {
        usedIds.add(picked.id);
        generated.push({ ...picked, orderKey: i });
      }
    }

    return generated;
  }

  function addStats(filter: string, count = 1) {
    const ticketType = getTicketType(filter);

    setStats((prev) => ({
      totalCombos: prev.totalCombos + count,
      normalTickets: prev.normalTickets + (ticketType === 'normal' ? count : 0),
      advancedTickets: prev.advancedTickets + (ticketType === 'advanced' ? count : 0),
    }));
  }

  function addHistory(generatedCards: CardData[], filter: string) {
    const ticketType = getTicketType(filter);

    setHistory((prev) => {
      const nextCardCounts = { ...prev.cardCounts };
      const nextNormalTeamCounts = { ...prev.normalTeamCounts };

      generatedCards.forEach((card) => {
        nextCardCounts[card.id] = (nextCardCounts[card.id] || 0) + 1;

        if (ticketType === 'normal') {
          const team = normalizeTeam(card.team);
          nextNormalTeamCounts[team] = (nextNormalTeamCounts[team] || 0) + 1;
        }
      });

      return {
        cardCounts: nextCardCounts,
        normalTeamCounts: nextNormalTeamCounts,
        comboOnlyCount: prev.comboOnlyCount + generatedCards.filter((card) => card.type === 'combo').length,
        wishHitCount: prev.wishHitCount + generatedCards.filter((card) => wishIds.includes(card.id)).length,
      };
    });
  }

  function resetTicketCounts() {
    if (!window.confirm('확정권 사용 횟수를 초기화하시겠습니까?')) return;

    setStats((prev) => ({
      ...prev,
      normalTickets: 0,
      advancedTickets: 0,
    }));
  }

  function toggleWish(cardId: string) {
    setWishIds((prev) =>
      prev.includes(cardId) ? prev.filter((id) => id !== cardId) : [...prev, cardId]
    );
  }

  function removeWishWithConfirm(cardId: string) {
    if (window.confirm('위시를 해제하시겠습니까?')) {
      setWishIds((prev) => prev.filter((id) => id !== cardId));
    }
  }

  function resetStats() {
    if (!window.confirm('시뮬레이션 카운터를 초기화하시겠습니까?')) return;
    setStats({ totalCombos: 0, normalTickets: 0, advancedTickets: 0 });
  }

  function resetHistory() {
    if (!window.confirm('획득 기록을 초기화하시겠습니까?')) return;
    setHistory({ cardCounts: {}, normalTeamCounts: {}, comboOnlyCount: 0, wishHitCount: 0 });
  }

  async function simulateCombo() {
    if (!dbLoaded || filteredNormalPool.length === 0) return;

    setIsRolling(true);
    setPickedCardId(null);
    setStage('ready');

    const generated = drawFive(comboMode, selectedFilter);

    setCards(generated);
    addStats(selectedFilter, 1);
    addHistory(generated, selectedFilter);
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

  function runAutoCombo() {
    if (!dbLoaded) return;

    setAutoResult('');
    setAutoResultItems([]);

    if (autoType === 'normal') {
      const allGenerated: CardData[] = [];

      for (let i = 0; i < autoCount; i++) {
        const generated = drawFive(autoMode, autoFilter);
        allGenerated.push(...generated);
      }

      if (allGenerated.length === 0) {
        setAutoResult('조건에 맞는 카드 풀이 없습니다.');
        return;
      }

      addStats(autoFilter, autoCount);
      addHistory(allGenerated, autoFilter);
      setAutoResult(`${autoMode === 'signature' ? '시그니처' : '임팩트'} 자동조합 ${autoCount}회 완료 / 총 ${allGenerated.length}장 획득`);
      setAutoResultItems(buildResultSummary(allGenerated).slice(0, 60));
      return;
    }

    const maxTry = 10000;
    let tries = 0;
    let foundCards: CardData[] = [];
    const search = specialSearch.trim().toLowerCase();

    while (tries < maxTry) {
      tries += 1;
      const generated = drawFive(autoMode, autoFilter);

      if (generated.length === 0) {
        setAutoResult('조건에 맞는 카드 풀이 없습니다.');
        return;
      }

      const matched =
        specialTarget === 'wish'
          ? generated.some((card) => wishIds.includes(card.id))
          : specialTarget === 'comboOnly'
            ? generated.some((card) => card.type === 'combo')
            : generated.some((card) =>
                [card.team, normalizeTeam(card.team), card.player, card.year, card.concept, card.position]
                  .join(' ')
                  .toLowerCase()
                  .includes(search)
              );

      if (matched) {
        foundCards = generated;
        break;
      }
    }

    if (foundCards.length === 0) {
      setAutoResult(`최대 ${maxTry.toLocaleString()}회까지 실행했지만 조건 카드가 등장하지 않았습니다. 특별 조합은 기록탭에 반영되지 않습니다.`);
      return;
    }

    setCards(foundCards);
    setStage('open');
    setPickedCardId(null);
    setMainTab('simulation');
    setAutoResult(`${tries.toLocaleString()}회 만에 조건 카드가 장판에 등장했습니다. 특별 조합은 기록탭에 반영되지 않았습니다.`);
    setAutoResultItems(buildResultSummary(foundCards));
  }

  function renderWishBadge(isWish: boolean) {
    if (!isWish) return null;

    return (
      <div className="absolute right-1 top-1 sm:right-2 sm:top-2 z-20 h-8 w-8 sm:h-10 sm:w-10 rounded-full bg-gradient-to-br from-blue-900 via-slate-800 to-black border-2 border-yellow-300 flex items-center justify-center shadow-[0_0_12px_rgba(250,204,21,0.85)]">
        <span className="text-yellow-300 text-lg sm:text-xl leading-none">★</span>
      </div>
    );
  }

  function renderCardMini(card: CardData, right?: React.ReactNode) {
    const isWish = wishIds.includes(card.id);

    return (
      <div className="relative flex items-center justify-between gap-3 rounded-2xl bg-zinc-900/80 border border-zinc-700 px-4 py-3">
        <div className="min-w-0">
          <div className="flex items-center gap-2">
            {isWish && <span className="text-yellow-300">★</span>}
            <div className="font-black text-white truncate">{getCardLabel(card)}</div>
          </div>
          <div className="text-xs text-zinc-500">
            {card.mode === 'signature' ? '시그니처' : '임팩트'} · {card.type === 'combo' ? '조합전용' : '일반'} · {normalizeTeam(card.team)}
          </div>
        </div>
        {right}
      </div>
    );
  }

  return (
    <div className="min-h-screen overflow-x-hidden bg-black text-white relative">
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
        <section className="w-full max-w-7xl flex flex-col gap-5">
          <div className="flex flex-col lg:flex-row items-center justify-between gap-4">
            <div className="text-center lg:text-left space-y-2">
              <h1 className="text-3xl md:text-5xl font-black tracking-wide text-pink-300 drop-shadow-[0_0_18px_rgba(255,105,180,0.6)]">
                컴프야V26 조합 시뮬레이터
              </h1>
              <p className="text-zinc-300">시그니처 · 임팩트 조합 시뮬레이션</p>
              <p className="text-sm text-zinc-500">
                {dbLoaded
                  ? `DB 로딩 완료 / 시그 ${signatureComboPool.length + signatureNormalPool.length}장 / 임팩 ${
                      impactComboPool.length + impactNormalPool.length
                    }장 / 위시 ${wishCards.length}장`
                  : '엑셀 DB 로딩 중...'}
              </p>
            </div>

            <section className="grid grid-cols-3 gap-2 rounded-3xl border border-white/10 bg-black/45 p-3 backdrop-blur">
              <div className="text-center px-3">
                <div className="text-xs text-zinc-500">총 조합</div>
                <div className="text-2xl font-black text-white">{stats.totalCombos}</div>
              </div>
              <div className="text-center px-3">
                <div className="mx-auto mb-1 h-10 w-8 rounded bg-gradient-to-b from-zinc-200 to-zinc-700 border border-zinc-400" />
                <div className="text-xs text-zinc-500">팀 일반 확정권</div>
                <div className="text-xl font-black text-zinc-200">{stats.normalTickets}</div>
              </div>
              <div className="text-center px-3">
                <div className="mx-auto mb-1 h-10 w-8 rounded bg-gradient-to-b from-yellow-300 to-zinc-900 border border-yellow-400" />
                <div className="text-xs text-zinc-500">팀 고급 확정권</div>
                <div className="text-xl font-black text-yellow-300">{stats.advancedTickets}</div>
              </div>

              <button
                onClick={resetTicketCounts}
                className="col-span-3 mt-1 rounded-xl bg-zinc-800 px-3 py-2 text-xs font-black text-zinc-300 hover:bg-zinc-700"
              >
                확정권 사용 초기화
              </button>
            </section>
          </div>

          <nav className="flex flex-wrap justify-center gap-3">
            {[
              ['simulation', '시뮬레이션'],
              ['wish', '위시'],
              ['history', '기록'],
            ].map(([tab, label]) => (
              <button
                key={tab}
                onClick={() => setMainTab(tab as MainTab)}
                className={`px-6 py-3 rounded-2xl font-black transition-all ${
                  mainTab === tab
                    ? 'bg-white text-black shadow-[0_0_20px_rgba(255,255,255,0.45)]'
                    : 'bg-zinc-900 text-zinc-300 border border-zinc-700 hover:bg-zinc-800'
                }`}
              >
                {label}
              </button>
            ))}
          </nav>
        </section>

        {mainTab === 'simulation' && (
          <>
            <section className="flex gap-3">
              <button
                onClick={() => {
                  setComboMode('signature');
                  setCards([]);
                  setStage('ready');
                }}
                className={`px-6 py-3 rounded-xl font-black transition-all ${
                  comboMode === 'signature'
                    ? 'bg-pink-500 text-white shadow-[0_0_18px_rgba(236,72,153,0.7)]'
                    : 'bg-zinc-800 text-zinc-300 hover:bg-zinc-700'
                }`}
              >
                시그니처
              </button>

              <button
                onClick={() => {
                  setComboMode('impact');
                  setCards([]);
                  setStage('ready');
                }}
                className={`px-6 py-3 rounded-xl font-black transition-all ${
                  comboMode === 'impact'
                    ? 'bg-lime-500 text-black shadow-[0_0_18px_rgba(132,204,22,0.7)]'
                    : 'bg-zinc-800 text-zinc-300 hover:bg-zinc-700'
                }`}
              >
                임팩트
              </button>
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
              onClick={simulateCombo}
              disabled={isRolling || !dbLoaded}
              className={`px-10 py-4 rounded-2xl text-xl font-black shadow-[0_0_25px_rgba(217,70,239,0.6)] hover:scale-105 transition-transform disabled:opacity-50 ${
                comboMode === 'signature'
                  ? 'bg-gradient-to-r from-pink-500 to-purple-600'
                  : 'bg-gradient-to-r from-lime-400 to-green-600 text-black'
              }`}
            >
              {isRolling ? '진행 중...' : comboMode === 'signature' ? '시그 조합 실행' : '임팩트 조합 실행'}
            </button>

            <button
              onClick={() => setAutoOpen((prev) => !prev)}
              className="px-7 py-3 rounded-2xl bg-zinc-900 border border-cyan-400/60 text-cyan-200 font-black shadow-[0_0_18px_rgba(34,211,238,0.25)] hover:bg-zinc-800"
            >
              자동 조합
            </button>

            {autoOpen && (
              <section className="w-full max-w-5xl rounded-3xl border border-cyan-400/25 bg-black/45 p-5 backdrop-blur space-y-5">
                <div className="flex flex-wrap gap-3 justify-center">
                  <button
                    onClick={() => setAutoType('normal')}
                    className={`px-5 py-2 rounded-xl font-black ${autoType === 'normal' ? 'bg-cyan-300 text-black' : 'bg-zinc-800 text-zinc-300'}`}
                  >
                    일반 자동조합
                  </button>
                  <button
                    onClick={() => setAutoType('special')}
                    className={`px-5 py-2 rounded-xl font-black ${autoType === 'special' ? 'bg-yellow-300 text-black' : 'bg-zinc-800 text-zinc-300'}`}
                  >
                    특별 조합
                  </button>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
                  <select value={autoMode} onChange={(e) => setAutoMode(e.target.value as ComboMode)} className="px-4 py-3 rounded-xl bg-zinc-900 border border-zinc-700">
                    <option value="signature">시그니처</option>
                    <option value="impact">임팩트</option>
                  </select>

                  <select value={autoFilter} onChange={(e) => setAutoFilter(e.target.value)} className="px-4 py-3 rounded-xl bg-zinc-900 border border-zinc-700">
                    {['전체', '드림', '나눔', ...allTeams].map((filter) => (
                      <option key={filter} value={filter}>
                        {filter}
                      </option>
                    ))}
                  </select>

                  {autoType === 'normal' ? (
                    <select value={autoCount} onChange={(e) => setAutoCount(Number(e.target.value))} className="px-4 py-3 rounded-xl bg-zinc-900 border border-zinc-700">
                      {autoCounts.map((count) => (
                        <option key={count} value={count}>
                          {count}회
                        </option>
                      ))}
                    </select>
                  ) : (
                    <select value={specialTarget} onChange={(e) => setSpecialTarget(e.target.value as SpecialTarget)} className="px-4 py-3 rounded-xl bg-zinc-900 border border-zinc-700">
                      <option value="wish">위시카드 등장까지</option>
                      <option value="comboOnly">조합 전용카드 등장까지</option>
                      <option value="specific">특정 카드 검색</option>
                    </select>
                  )}

                  <button onClick={runAutoCombo} disabled={!dbLoaded} className="px-5 py-3 rounded-xl bg-cyan-400 text-black font-black disabled:opacity-40">
                    실행
                  </button>
                </div>

                {autoType === 'special' && specialTarget === 'specific' && (
                  <input
                    value={specialSearch}
                    onChange={(e) => setSpecialSearch(e.target.value)}
                    placeholder="선수명, 팀, 시즌, 컨셉, 포지션 검색"
                    className="w-full px-4 py-3 rounded-xl bg-zinc-900 border border-zinc-700 outline-none focus:border-yellow-300"
                  />
                )}

                {autoType === 'special' && (
                  <p className="text-sm text-yellow-300">
                    특별 조합은 원하는 카드가 장판에 등장할 때까지 빠르게 실행하며, 기록탭에는 반영되지 않습니다.
                  </p>
                )}

                {autoResult && (
                  <div className="rounded-2xl bg-zinc-950 border border-zinc-700 p-4">
                    <div className="font-black text-cyan-200 mb-3">{autoResult}</div>
                    {autoResultItems.length > 0 && (
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-2 max-h-72 overflow-y-auto pr-1">
                        {autoResultItems.map(({ card, count }) => {
                          const special = card.type === 'combo' || wishIds.includes(card.id);
                          return (
                            <div key={card.id} className={`rounded-xl px-3 py-2 text-sm ${special ? 'bg-yellow-300 text-black font-black' : 'bg-zinc-900 text-zinc-300'}`}>
                              {wishIds.includes(card.id) ? '★ ' : ''}
                              {card.type === 'combo' ? '[조합전용] ' : ''}
                              {getCardLabel(card)} × {count}
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                )}
              </section>
            )}

            {stage === 'open' && (
              <button
                onClick={startShuffle}
                className="px-8 py-3 rounded-2xl bg-gradient-to-r from-yellow-400 to-orange-500 text-black font-black shadow-[0_0_20px_rgba(255,180,0,0.7)] hover:scale-105 transition-transform"
              >
                셔플 시작
              </button>
            )}

            <p className="h-6 text-sm text-zinc-400">
              {stage === 'open' && '카드 5장이 공개되었습니다. 위시 카드는 오른쪽 상단에 별이 표시됩니다.'}
              {stage === 'back' && '카드를 뒤집는 중...'}
              {stage === 'shuffling' && '카드를 섞는 중...'}
              {stage === 'shuffled' && '뒷면 카드 1장을 선택하세요'}
              {stage === 'picked' && '선택 결과 공개'}
            </p>

            <section className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-3 sm:gap-5 lg:gap-6 max-w-7xl min-h-96">
              {cards.map((card) => {
                const isVisible = cardFaceVisible(card.id);
                const isPicked = pickedCardId === card.id;
                const isImpact = card.mode === 'impact';
                const isWish = wishIds.includes(card.id);

                return (
                  <button
                    key={card.id}
                    onClick={() => pickCard(card.id)}
                    disabled={stage !== 'shuffled'}
                    className={`relative w-24 h-36 sm:w-32 sm:h-48 md:w-40 md:h-60 lg:w-52 lg:h-80
                      rounded-2xl sm:rounded-3xl transition-all duration-500 [perspective:1000px]
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
                        className={`absolute inset-0 rounded-2xl sm:rounded-3xl overflow-hidden border-2 sm:border-4 [backface-visibility:hidden] ${
                          card.type === 'combo'
                            ? 'border-orange-400 shadow-[0_0_45px_rgba(255,140,0,0.95)]'
                            : isImpact
                              ? 'border-lime-300 shadow-[0_0_22px_rgba(132,204,22,0.6)]'
                              : 'border-pink-300 shadow-[0_0_22px_rgba(255,105,180,0.55)]'
                        }`}
                      >
                        {renderWishBadge(isWish)}

                        <div
                          className={`absolute inset-0 ${
                            isImpact
                              ? 'bg-gradient-to-br from-lime-300 via-lime-500 to-black'
                              : 'bg-gradient-to-b from-pink-200 via-pink-400 to-pink-100'
                          }`}
                        />
                        <div className="absolute inset-0 bg-[radial-gradient(circle_at_center,rgba(255,255,255,0.55),transparent_55%)]" />
                        <div className="absolute inset-x-0 bottom-0 h-16 sm:h-20 lg:h-24 bg-white/40" />

                        {card.type === 'combo' && <div className="absolute inset-0 animate-pulse bg-orange-400/10" />}

                        <div className="relative z-10 flex h-full flex-col p-2 sm:p-3 lg:p-4">
                          <div className="flex items-start justify-between">
                            <div className={`text-2xl sm:text-3xl lg:text-4xl font-black leading-none ${isImpact ? 'text-lime-100' : 'text-pink-700'}`}>
                              {card.position}
                            </div>

                            <div className="h-8 w-8 sm:h-10 sm:w-10 lg:h-12 lg:w-12 rounded-full bg-white/95 flex items-center justify-center shadow-lg p-1 overflow-hidden">
                              <img src={`/logos/${getLogoTeam(card.team)}.png`} alt={card.team} className="h-full w-full object-contain" />
                            </div>
                          </div>

                          <div className="flex-1 flex items-center justify-center">
                            <div className="text-2xl sm:text-3xl lg:text-5xl font-black text-white/30 italic select-none">
                              {isImpact ? 'Impact' : 'Signature'}
                            </div>
                          </div>

                          <div className="text-center pb-1">
                            {card.type === 'combo' && (
                              <div className="text-orange-600 font-black text-[10px] sm:text-xs lg:text-sm tracking-widest mb-1">
                                조합전용
                              </div>
                            )}

                            <div className="text-lg sm:text-2xl lg:text-3xl font-black text-zinc-950 tracking-tight">
                              {card.player}
                            </div>

                            <div className="text-sm sm:text-lg lg:text-xl font-black text-zinc-700">
                              {isImpact ? card.concept : `'${card.year}`}
                            </div>
                          </div>
                        </div>
                      </div>

                      <div className="absolute inset-0 rounded-2xl sm:rounded-3xl overflow-hidden border-2 sm:border-4 border-fuchsia-300 bg-gradient-to-br from-fuchsia-700 via-pink-500 to-purple-800 shadow-[0_0_30px_rgba(217,70,239,0.7)] [backface-visibility:hidden] [transform:rotateY(180deg)]">
                        <div className="absolute inset-2 sm:inset-3 rounded-2xl border border-white/30" />
                        <div className="absolute inset-0 bg-[radial-gradient(circle_at_center,rgba(255,255,255,0.25),transparent_55%)]" />
                        <div className="relative z-10 h-full flex items-center justify-center">
                          <div className="text-2xl sm:text-3xl lg:text-4xl font-black text-white/90 italic tracking-wider">
                            {isImpact ? 'IMPACT' : 'SIGN'}
                          </div>
                        </div>
                      </div>
                    </div>
                  </button>
                );
              })}
            </section>
          </>
        )}

        {mainTab === 'wish' && (
          <section className="w-full max-w-6xl rounded-3xl border border-white/10 bg-black/40 p-5 backdrop-blur space-y-5">
            <div className="flex flex-wrap justify-center gap-3">
              <button
                onClick={() => setWishSubTab('manage')}
                className={`px-6 py-3 rounded-2xl font-black ${wishSubTab === 'manage' ? 'bg-yellow-300 text-black' : 'bg-zinc-800 text-zinc-300'}`}
              >
                위시 관리
              </button>
              <button
                onClick={() => setWishSubTab('register')}
                className={`px-6 py-3 rounded-2xl font-black ${wishSubTab === 'register' ? 'bg-yellow-300 text-black' : 'bg-zinc-800 text-zinc-300'}`}
              >
                위시 등록
              </button>
            </div>

            {wishSubTab === 'manage' && (
              <>
                <h2 className="text-2xl font-black text-yellow-300">위시 관리</h2>
                {wishCards.length === 0 ? (
                  <p className="text-zinc-500">아직 등록된 위시 카드가 없습니다.</p>
                ) : (
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    {wishCards.map((card) =>
                      renderCardMini(
                        card,
                        <button onClick={() => removeWishWithConfirm(card.id)} className="shrink-0 px-3 py-2 rounded-xl bg-red-500 text-white text-xs font-black">
                          위시 해제
                        </button>
                      )
                    )}
                  </div>
                )}
              </>
            )}

            {wishSubTab === 'register' && (
              <>
                <h2 className="text-2xl font-black text-yellow-300">위시 등록</h2>
                <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
                  <input
                    value={wishSearch}
                    onChange={(e) => setWishSearch(e.target.value)}
                    placeholder="선수명, 컨셉, 팀, 시즌 검색"
                    className="md:col-span-4 px-4 py-3 rounded-xl bg-zinc-900 border border-zinc-700 outline-none focus:border-yellow-300"
                  />
                  <select value={wishModeFilter} onChange={(e) => setWishModeFilter(e.target.value as 'all' | ComboMode)} className="px-4 py-3 rounded-xl bg-zinc-900 border border-zinc-700">
                    <option value="all">전체 타입</option>
                    <option value="signature">시그니처</option>
                    <option value="impact">임팩트</option>
                  </select>
                  <select value={wishTypeFilter} onChange={(e) => setWishTypeFilter(e.target.value as 'all' | 'combo' | 'normal')} className="px-4 py-3 rounded-xl bg-zinc-900 border border-zinc-700">
                    <option value="all">전체 카드</option>
                    <option value="combo">조합전용카드만</option>
                    <option value="normal">일반카드만</option>
                  </select>
                  <select value={wishTeamFilter} onChange={(e) => setWishTeamFilter(e.target.value)} className="px-4 py-3 rounded-xl bg-zinc-900 border border-zinc-700">
                    {['전체', ...allTeams].map((team) => (
                      <option key={team} value={team}>
                        {team}
                      </option>
                    ))}
                  </select>
                  <div className="px-4 py-3 rounded-xl bg-zinc-950 border border-zinc-800 text-zinc-400">
                    검색 {wishSearchResults.length}개
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-3 max-h-[520px] overflow-y-auto pr-1">
                  {wishSearchResults.map((card) => {
                    const isWish = wishIds.includes(card.id);

                    return renderCardMini(
                      card,
                      <button
                        onClick={() => (isWish ? removeWishWithConfirm(card.id) : toggleWish(card.id))}
                        className={`shrink-0 px-3 py-2 rounded-xl text-xs font-black ${
                          isWish ? 'bg-yellow-300 text-black' : 'bg-zinc-800 text-yellow-300 border border-yellow-300/40'
                        }`}
                      >
                        {isWish ? '등록됨' : '위시 등록'}
                      </button>
                    );
                  })}
                </div>
              </>
            )}
          </section>
        )}

        {mainTab === 'history' && (
          <section className="w-full max-w-6xl rounded-3xl border border-white/10 bg-black/40 p-5 backdrop-blur space-y-5">
            <div className="flex items-center justify-between gap-3">
              <h2 className="text-2xl font-black text-cyan-200">기록</h2>
              <div className="flex gap-2">
                <button onClick={resetStats} className="px-4 py-2 rounded-xl bg-zinc-800 text-zinc-300 text-sm font-bold">
                  카운터 초기화
                </button>
                <button onClick={resetHistory} className="px-4 py-2 rounded-xl bg-red-500/80 text-white text-sm font-bold">
                  기록 초기화
                </button>
              </div>
            </div>

            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              <div className="rounded-2xl bg-zinc-900 border border-zinc-700 p-4">
                <div className="text-sm text-zinc-500">총 획득 카드</div>
                <div className="text-3xl font-black">{Object.values(history.cardCounts).reduce((a, b) => a + b, 0)}</div>
              </div>
              <div className="rounded-2xl bg-zinc-900 border border-zinc-700 p-4">
                <div className="text-sm text-zinc-500">조합 전용카드 획득</div>
                <div className="text-3xl font-black text-orange-300">{history.comboOnlyCount}</div>
              </div>
              <div className="rounded-2xl bg-zinc-900 border border-zinc-700 p-4">
                <div className="text-sm text-zinc-500">위시카드 획득</div>
                <div className="text-3xl font-black text-yellow-300">{history.wishHitCount}</div>
              </div>
              <div className="rounded-2xl bg-zinc-900 border border-zinc-700 p-4">
                <div className="text-sm text-zinc-500">기록된 카드 종류</div>
                <div className="text-3xl font-black text-cyan-200">{historyCards.length}</div>
              </div>
            </div>

            <div className="rounded-2xl bg-zinc-950 border border-zinc-800 p-4">
              <h3 className="font-black mb-3 text-pink-300">일반 확정권 사용 시 팀별 획득</h3>
              <div className="grid grid-cols-2 md:grid-cols-5 gap-2">
                {allTeams.map((team) => (
                  <div key={team} className="rounded-xl bg-zinc-900 px-3 py-2 text-sm flex justify-between">
                    <span>{team}</span>
                    <b>{history.normalTeamCounts[team] || 0}</b>
                  </div>
                ))}
              </div>
            </div>

            <div className="rounded-2xl bg-zinc-950 border border-zinc-800 p-4">
              <h3 className="font-black mb-3 text-cyan-200">카드별 획득 횟수</h3>
              {historyCards.length === 0 ? (
                <p className="text-zinc-500">아직 획득 기록이 없습니다.</p>
              ) : (
                <div className="grid grid-cols-1 md:grid-cols-2 gap-3 max-h-[520px] overflow-y-auto pr-1">
                  {historyCards.map(({ card, count }) =>
                    renderCardMini(
                      card,
                      <div className="shrink-0 rounded-xl bg-cyan-300 px-3 py-2 text-sm font-black text-black">
                        × {count}
                      </div>
                    )
                  )}
                </div>
              )}
            </div>
          </section>
        )}
      </main>
    </div>
  );
}
