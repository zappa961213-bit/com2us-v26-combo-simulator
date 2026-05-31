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
type ComboMode = 'signature' | 'impact' | 'custom';
type MainTab = 'simulation' | 'wish';
type InfoTab = 'intro' | 'guide' | 'updates' | 'contact';
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

function getPlayerNameTextSize(player: string) {
  const koreanLength = Array.from(player).filter((char) => /[가-힣]/.test(char)).length;

  if (koreanLength >= 6) {
    return 'text-sm sm:text-xl lg:text-2xl';
  }

  if (koreanLength >= 5) {
    return 'text-[15px] sm:text-xl lg:text-2xl';
  }

  return 'text-lg sm:text-2xl lg:text-3xl';
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

function safeStringArray(value: unknown) {
  return Array.isArray(value) ? value.filter((item): item is string => typeof item === 'string') : [];
}

function safeSimStats(value: unknown): SimStats {
  if (!value || typeof value !== 'object') {
    return { totalCombos: 0, normalTickets: 0, advancedTickets: 0 };
  }

  const stats = value as Partial<SimStats>;

  return {
    totalCombos: typeof stats.totalCombos === 'number' ? stats.totalCombos : 0,
    normalTickets: typeof stats.normalTickets === 'number' ? stats.normalTickets : 0,
    advancedTickets: typeof stats.advancedTickets === 'number' ? stats.advancedTickets : 0,
  };
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
  const [infoTab, setInfoTab] = useState<InfoTab>('intro');
  const [wishSubTab, setWishSubTab] = useState<WishSubTab>('manage');

  const [comboMode, setComboMode] = useState<ComboMode>('signature');
  const [customSlots, setCustomSlots] = useState<(CardData | null)[]>([null, null, null, null, null]);
  const [customPickerIndex, setCustomPickerIndex] = useState<number | null>(null);
  const [customSearch, setCustomSearch] = useState('');
  const [customModeFilter, setCustomModeFilter] = useState<'all' | 'signature' | 'impact'>('all');
  const [customTypeFilter, setCustomTypeFilter] = useState<'all' | 'combo' | 'normal'>('all');
  const [customTeamFilter, setCustomTeamFilter] = useState('전체');
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
    const savedWishIds = safeJsonParse<unknown>(localStorage.getItem('wishIds'), []);
    const savedStats = safeJsonParse<unknown>(localStorage.getItem('simStats'), {
      totalCombos: 0,
      normalTickets: 0,
      advancedTickets: 0,
    });

    setWishIds(safeStringArray(savedWishIds));
    setStats(safeSimStats(savedStats));
  }, []);

  useEffect(() => {
    localStorage.setItem('wishIds', JSON.stringify(wishIds));
  }, [wishIds]);

  useEffect(() => {
    localStorage.setItem('simStats', JSON.stringify(stats));
  }, [stats]);

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

  const customSearchResults = useMemo(() => {
    const keyword = customSearch.trim().toLowerCase();

    return allPool
      .filter((card) => {
        if (customModeFilter !== 'all' && card.mode !== customModeFilter) return false;
        if (customTypeFilter !== 'all' && card.type !== customTypeFilter) return false;
        if (customTeamFilter !== '전체' && normalizeTeam(card.team) !== customTeamFilter) return false;

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
  }, [allPool, customSearch, customModeFilter, customTypeFilter, customTeamFilter]);

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

  function drawFive(mode: ComboMode, filter: string) {
    const normalPool = mode === 'signature' ? signatureNormalPool : impactNormalPool;
    const comboPool = mode === 'signature' ? signatureComboPool : impactComboPool;

    const filteredNormal = normalPool.filter((card) => matchesFilter(card, filter));
    const filteredCombo = comboPool.filter((card) => matchesFilter(card, filter));

    const filteredSignatureNormal = signatureNormalPool.filter((card) => matchesFilter(card, filter));
    const filteredSignatureCombo = signatureComboPool.filter((card) => matchesFilter(card, filter));

    if (filteredNormal.length === 0 && filteredCombo.length === 0) return [];

    const generated: CardData[] = [];
    const usedIds = new Set<string>();

    for (let i = 0; i < 5; i++) {
      let pool: CardData[] = filteredNormal;

      if (mode === 'impact') {
        const useSignatureBoardCard =
          Math.random() < 0.015 && (filteredSignatureNormal.length > 0 || filteredSignatureCombo.length > 0);

        if (useSignatureBoardCard) {
          const useSignatureCombo = Math.random() < 0.09;
          pool =
            useSignatureCombo && filteredSignatureCombo.length > 0
              ? filteredSignatureCombo
              : filteredSignatureNormal.length > 0
                ? filteredSignatureNormal
                : filteredSignatureCombo;
        } else {
          const useImpactCombo = Math.random() < 0.15;
          pool = useImpactCombo && filteredCombo.length > 0 ? filteredCombo : filteredNormal;
        }
      } else {
        const useSignatureCombo = Math.random() < 0.09;
        pool = useSignatureCombo && filteredCombo.length > 0 ? filteredCombo : filteredNormal;
      }

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

  function pickCustomCard(card: CardData) {
    if (customPickerIndex === null) return;

    setCustomSlots((prev) => {
      const next = [...prev];
      next[customPickerIndex] = { ...card, orderKey: customPickerIndex };
      return next;
    });
    setCustomPickerIndex(null);
    setCustomSearch('');
  }

  function clearCustomSlot(index: number) {
    setCustomSlots((prev) => {
      const next = [...prev];
      next[index] = null;
      return next;
    });
  }

  function applyCustomBoard() {
    const filledCards = customSlots.filter((card): card is CardData => Boolean(card));

    if (filledCards.length !== 5) {
      window.alert('커스텀 장판 5칸을 모두 설정해주세요.');
      return;
    }

    setPickedCardId(null);
    setCards(filledCards.map((card, index) => ({ ...card, orderKey: index })));
    setStage('open');
    setMainTab('simulation');
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

  function resetTicketCounts() {
    if (!window.confirm('확정권 사용 횟수를 초기화하시겠습니까?')) return;

    setStats((prev) => ({
      ...prev,
      normalTickets: 0,
      advancedTickets: 0,
    }));
  }

  async function simulateCombo() {
    if (comboMode === 'custom') {
      applyCustomBoard();
      return;
    }

    if (!dbLoaded || filteredNormalPool.length === 0) return;

    setIsRolling(true);
    setPickedCardId(null);
    setStage('ready');

    const generated = drawFive(comboMode, selectedFilter);

    setCards(generated);
    addStats(selectedFilter, 1);
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
      const acquiredCards: CardData[] = [];
      let boardCount = 0;

      for (let i = 0; i < autoCount; i++) {
        const boardCards = drawFive(autoMode, autoFilter);

        if (boardCards.length > 0) {
          boardCount += boardCards.length;

          // 자동조합 1회 = 장판 5장 등장 후 그중 1장 획득
          const acquiredCard = boardCards[Math.floor(Math.random() * boardCards.length)];
          acquiredCards.push(acquiredCard);
        }
      }

      if (acquiredCards.length === 0) {
        setAutoResult('조건에 맞는 카드 풀이 없습니다.');
        return;
      }

      const comboOnlyHits = acquiredCards.filter((card) => card.type === 'combo').length;
      const wishHits = acquiredCards.filter((card) => wishIds.includes(card.id)).length;

      addStats(autoFilter, acquiredCards.length);
      setAutoResult(
        `${autoMode === 'signature' ? '시그니처' : '임팩트'} 자동조합 ${acquiredCards.length}회 완료 / ` +
          `장판 ${boardCount}장 확인 / 최종 ${acquiredCards.length}장 획득` +
          ` / 조합전용 ${comboOnlyHits}장 / 위시 ${wishHits}장`
      );
      setAutoResultItems(buildResultSummary(acquiredCards).slice(0, 60));
      return;
    }

    const maxTry = 10000;
    let tries = 0;
    let foundCards: CardData[] = [];
    const search = specialSearch.trim().toLowerCase();

    if (specialTarget === 'specific' && !search) {
      setAutoResult('특정 카드 검색어를 입력해주세요.');
      return;
    }

    while (tries < maxTry) {
      tries += 1;
      const boardCards = drawFive(autoMode, autoFilter);

      if (boardCards.length === 0) {
        setAutoResult('조건에 맞는 카드 풀이 없습니다.');
        return;
      }

      const matched =
        specialTarget === 'wish'
          ? boardCards.some((card) => wishIds.includes(card.id))
          : specialTarget === 'comboOnly'
            ? boardCards.some((card) => card.type === 'combo')
            : boardCards.some((card) =>
                [card.team, normalizeTeam(card.team), card.player, card.year, card.concept, card.position]
                  .join(' ')
                  .toLowerCase()
                  .includes(search)
              );

      if (matched) {
        foundCards = boardCards;
        break;
      }
    }

    if (foundCards.length === 0) {
      setAutoResult(`최대 ${maxTry.toLocaleString()}회까지 실행했지만 조건 카드가 장판에 등장하지 않았습니다.`);
      return;
    }

    const targetLabel =
      specialTarget === 'wish'
        ? '위시카드'
        : specialTarget === 'comboOnly'
          ? '조합 전용카드'
          : `검색어 "${specialSearch.trim()}"`;

    const matchedCards = foundCards.filter((card) => {
      if (specialTarget === 'wish') return wishIds.includes(card.id);
      if (specialTarget === 'comboOnly') return card.type === 'combo';

      return [card.team, normalizeTeam(card.team), card.player, card.year, card.concept, card.position]
        .join(' ')
        .toLowerCase()
        .includes(search);
    });

    setCards(foundCards);
    setStage('open');
    setPickedCardId(null);
    setMainTab('simulation');
    setAutoResult(
      `${targetLabel}가 ${tries.toLocaleString()}회 만에 장판에 등장했습니다. ` +
        `총 ${tries * 5}장의 장판 카드를 확인했습니다.`
    );
    setAutoResultItems(buildResultSummary(matchedCards.length > 0 ? matchedCards : foundCards));
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

      <main className="relative z-10 flex flex-col items-center px-4 py-5 md:py-8 gap-4 md:gap-7">
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
            <section className="flex gap-2 sm:gap-3">
              {[
                ['signature', '시그니처'],
                ['impact', '임팩트'],
                ['custom', '커스텀'],
              ].map(([mode, label]) => (
                <button
                  key={mode}
                  onClick={() => {
                    setComboMode(mode as ComboMode);
                    setCards([]);
                    setStage('ready');
                    setPickedCardId(null);
                  }}
                  className={`px-5 sm:px-6 py-3 rounded-xl font-black transition-all ${
                    comboMode === mode
                      ? mode === 'impact'
                        ? 'bg-lime-500 text-black shadow-[0_0_18px_rgba(132,204,22,0.7)]'
                        : mode === 'custom'
                          ? 'bg-cyan-300 text-black shadow-[0_0_18px_rgba(103,232,249,0.7)]'
                          : 'bg-pink-500 text-white shadow-[0_0_18px_rgba(236,72,153,0.7)]'
                      : 'bg-zinc-800 text-zinc-300 hover:bg-zinc-700'
                  }`}
                >
                  {label}
                </button>
              ))}
            </section>

            {comboMode !== 'custom' && (
              <>
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

              </>
            )}

            <section className="flex flex-col items-center gap-2 w-full max-w-[320px] sm:max-w-md">
              <button
                onClick={simulateCombo}
                disabled={isRolling || !dbLoaded}
                className={`w-full px-6 py-3 sm:px-10 sm:py-4 rounded-2xl text-lg sm:text-xl font-black shadow-[0_0_25px_rgba(217,70,239,0.6)] hover:scale-105 transition-transform disabled:opacity-50 ${
                  comboMode === 'signature'
                    ? 'bg-gradient-to-r from-pink-500 to-purple-600'
                    : 'bg-gradient-to-r from-lime-400 to-green-600 text-black'
                }`}
              >
                {isRolling ? '진행 중...' : comboMode === 'custom' ? '커스텀 장판 적용' : comboMode === 'signature' ? '시그 조합 실행' : '임팩트 조합 실행'}
              </button>

              <div className="grid grid-cols-2 gap-2 w-full">
                <button
                  onClick={() => setAutoOpen((prev) => !prev)}
                  disabled={comboMode === 'custom'}
                  className="px-3 py-2.5 sm:px-4 sm:py-3 rounded-2xl bg-zinc-900 border border-cyan-400/60 text-cyan-200 text-sm sm:text-base font-black shadow-[0_0_18px_rgba(34,211,238,0.25)] hover:bg-zinc-800 disabled:opacity-30 disabled:grayscale"
                >
                  자동 조합
                </button>

                <button
                  onClick={startShuffle}
                  disabled={stage !== 'open'}
                  className="px-3 py-2.5 sm:px-4 sm:py-3 rounded-2xl bg-gradient-to-r from-yellow-400 to-orange-500 text-black text-sm sm:text-base font-black shadow-[0_0_20px_rgba(255,180,0,0.7)] hover:scale-105 transition-transform disabled:opacity-30 disabled:grayscale"
                >
                  셔플 시작
                </button>
              </div>
            </section>

            {comboMode === 'custom' && (
              <section className="w-full max-w-5xl rounded-3xl border border-cyan-400/25 bg-black/45 p-5 backdrop-blur space-y-5">
                <div>
                  <h2 className="text-xl font-black text-cyan-200">커스텀 조합</h2>
                  <p className="mt-2 text-sm text-zinc-400">
                    장판 5칸을 직접 설정한 뒤 커스텀 장판 적용을 누르면 기존 셔플 방식으로 1장을 획득할 수 있습니다.
                  </p>
                </div>

                <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
                  {customSlots.map((slot, index) => (
                    <div key={index} className="rounded-2xl border border-zinc-700 bg-zinc-950/80 p-3 text-center space-y-3">
                      <div className="text-xs font-black text-zinc-500">{index + 1}번 장판</div>

                      {slot ? (
                        <div className="space-y-2">
                          <div className="font-black text-white leading-tight">
                            {slot.type === 'combo' ? '[조합전용] ' : ''}
                            {getCardLabel(slot)}
                          </div>
                          <div className="text-xs text-zinc-500">
                            {slot.mode === 'signature' ? '시그니처' : '임팩트'} · {slot.type === 'combo' ? '조합전용' : '일반'}
                          </div>
                          <button
                            onClick={() => clearCustomSlot(index)}
                            className="w-full rounded-xl bg-red-500/80 px-3 py-2 text-xs font-black text-white hover:bg-red-500"
                          >
                            비우기
                          </button>
                        </div>
                      ) : (
                        <div className="text-sm text-zinc-500">선수 미설정</div>
                      )}

                      <button
                        onClick={() => setCustomPickerIndex(index)}
                        className="w-full rounded-xl bg-cyan-300 px-3 py-2 text-sm font-black text-black hover:bg-cyan-200"
                      >
                        선수 설정
                      </button>
                    </div>
                  ))}
                </div>

                {customPickerIndex !== null && (
                  <div className="rounded-3xl border border-white/10 bg-zinc-950 p-4 space-y-4">
                    <div className="flex items-center justify-between gap-3">
                      <h3 className="font-black text-cyan-200">{customPickerIndex + 1}번 장판 선수 설정</h3>
                      <button
                        onClick={() => setCustomPickerIndex(null)}
                        className="rounded-xl bg-zinc-800 px-3 py-2 text-xs font-black text-zinc-300 hover:bg-zinc-700"
                      >
                        닫기
                      </button>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
                      <input
                        value={customSearch}
                        onChange={(e) => setCustomSearch(e.target.value)}
                        placeholder="선수명, 컨셉, 팀, 시즌 검색"
                        className="md:col-span-4 px-4 py-3 rounded-xl bg-zinc-900 border border-zinc-700 outline-none focus:border-cyan-300"
                      />
                      <select value={customModeFilter} onChange={(e) => setCustomModeFilter(e.target.value as 'all' | 'signature' | 'impact')} className="px-4 py-3 rounded-xl bg-zinc-900 border border-zinc-700">
                        <option value="all">전체 타입</option>
                        <option value="signature">시그니처</option>
                        <option value="impact">임팩트</option>
                      </select>
                      <select value={customTypeFilter} onChange={(e) => setCustomTypeFilter(e.target.value as 'all' | 'combo' | 'normal')} className="px-4 py-3 rounded-xl bg-zinc-900 border border-zinc-700">
                        <option value="all">전체 카드</option>
                        <option value="combo">조합전용카드만</option>
                        <option value="normal">일반카드만</option>
                      </select>
                      <select value={customTeamFilter} onChange={(e) => setCustomTeamFilter(e.target.value)} className="px-4 py-3 rounded-xl bg-zinc-900 border border-zinc-700">
                        {['전체', ...allTeams].map((team) => (
                          <option key={team} value={team}>
                            {team}
                          </option>
                        ))}
                      </select>
                      <div className="px-4 py-3 rounded-xl bg-zinc-950 border border-zinc-800 text-zinc-400">
                        검색 {customSearchResults.length}개
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-3 max-h-80 overflow-y-auto pr-1">
                      {customSearchResults.map((card) => {
                        const isWish = wishIds.includes(card.id);

                        return (
                          <button
                            key={card.id}
                            onClick={() => pickCustomCard(card)}
                            className="text-left rounded-2xl bg-zinc-900/80 border border-zinc-700 px-4 py-3 hover:border-cyan-300"
                          >
                            <div className="flex items-center gap-2">
                              {isWish && <span className="text-yellow-300">★</span>}
                              <div className="font-black text-white">
                                {card.type === 'combo' ? '[조합전용] ' : ''}
                                {getCardLabel(card)}
                              </div>
                            </div>
                            <div className="text-xs text-zinc-500">
                              {card.mode === 'signature' ? '시그니처' : '임팩트'} · {card.type === 'combo' ? '조합전용' : '일반'} · {normalizeTeam(card.team)}
                            </div>
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}
              </section>
            )}

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
                    특별 조합은 조건 카드가 장판에 등장할 때까지 빠르게 시뮬레이션하며, 몇 회 만에 등장했는지 표시합니다.
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

            <p className="min-h-5 text-xs sm:text-sm text-zinc-400 text-center px-2">
              {stage === 'open' && (comboMode === 'custom' ? '커스텀 장판이 적용되었습니다. 셔플을 시작하세요.' : '카드 5장이 공개되었습니다. 셔플을 시작하세요.')}
              {stage === 'back' && '카드를 뒤집는 중...'}
              {stage === 'shuffling' && '카드를 섞는 중...'}
              {stage === 'shuffled' && '뒷면 카드 1장을 선택하세요'}
              {stage === 'picked' && '선택 결과 공개'}
            </p>

            <section className="grid grid-cols-6 gap-3 sm:gap-5 lg:gap-6 max-w-5xl min-h-96 justify-items-center">
              {[...cards]
                .sort((a, b) => a.orderKey - b.orderKey)
                .map((card, displayIndex) => {
                const isVisible = cardFaceVisible(card.id);
                const isPicked = pickedCardId === card.id;
                const isImpact = card.mode === 'impact';
                const isWish = wishIds.includes(card.id);

                return (
                  <button
                    key={card.id}
                    onClick={() => pickCard(card.id)}
                    disabled={stage !== 'shuffled'}
                    className={`relative col-span-2 ${displayIndex === 3 ? 'col-start-2' : displayIndex === 4 ? 'col-start-4' : ''}
                      w-24 h-36 sm:w-32 sm:h-48 md:w-40 md:h-60 lg:w-52 lg:h-80
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
                  >
                    <div
                      className={`relative w-full h-full transition-transform duration-700 [transform-style:preserve-3d] [-webkit-transform-style:preserve-3d] [will-change:transform] ${
                        isVisible ? '[transform:rotateY(0deg)]' : '[transform:rotateY(180deg)]'
                      }`}
                    >
                      <div
                        className={`absolute inset-0 rounded-2xl sm:rounded-3xl overflow-hidden border-2 sm:border-4 transition-opacity duration-100 [backface-visibility:hidden] [-webkit-backface-visibility:hidden] ${!isVisible ? 'opacity-0 pointer-events-none' : 'opacity-100'} ${
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

                            <div className={`${getPlayerNameTextSize(card.player)} font-black text-zinc-950 tracking-tight leading-none sm:leading-tight`}>
                              {card.player}
                            </div>

                            <div className="text-sm sm:text-lg lg:text-xl font-black text-zinc-700">
                              {isImpact ? card.concept : `'${card.year}`}
                            </div>
                          </div>
                        </div>
                      </div>

                      <div
                        className={`absolute inset-0 rounded-2xl sm:rounded-3xl overflow-hidden border-2 sm:border-4 border-zinc-400 bg-gradient-to-br from-zinc-200 via-zinc-500 to-zinc-900 shadow-[0_0_24px_rgba(161,161,170,0.45)] transition-opacity duration-100 [backface-visibility:hidden] [-webkit-backface-visibility:hidden] [transform:rotateY(180deg)] ${isVisible ? 'opacity-0 pointer-events-none' : 'opacity-100'}`}
                      >
                        <div className="absolute inset-2 sm:inset-3 rounded-2xl border border-white/25" />
                        <div className="absolute inset-0 bg-[linear-gradient(135deg,rgba(255,255,255,0.28),transparent_35%,rgba(0,0,0,0.22))]" />
                        <div className="absolute inset-0 bg-[radial-gradient(circle_at_top_left,rgba(255,255,255,0.35),transparent_45%)]" />
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


        <section className="w-full max-w-5xl mt-8 rounded-3xl border border-white/10 bg-white/90 text-slate-800 p-5 sm:p-6 shadow-xl">
          <div className="mb-5 flex flex-wrap justify-center gap-4 text-sm font-black text-slate-600">
            {[
              ['intro', '소개'],
              ['guide', '사용 가이드'],
              ['updates', '업데이트 내역'],
              ['contact', '문의'],
            ].map(([tab, label]) => (
              <button
                key={tab}
                onClick={() => setInfoTab(tab as InfoTab)}
                className={`transition-colors ${
                  infoTab === tab
                    ? 'text-pink-600 underline underline-offset-4'
                    : 'hover:text-slate-950'
                }`}
              >
                {label}
              </button>
            ))}
          </div>

          {infoTab === 'intro' && (
            <div className="space-y-5">
              <section>
                <h2 className="text-xl font-black text-slate-900 mb-3">소개</h2>
                <p className="leading-7 font-semibold">
                  이 사이트는 컴프야V26 카드 조합 결과를 가볍게 시뮬레이션해볼 수 있는 비공식 팬메이드 도구입니다.
                  시그니처, 임팩트 조합과 팀 확정권 조건을 바탕으로 조합 결과를 확인할 수 있습니다.
                </p>
              </section>

              <hr className="border-slate-300" />

              <section>
                <h3 className="text-lg font-black text-slate-900 mb-3">제공하는 기능</h3>
                <ul className="list-disc pl-5 space-y-2 font-semibold leading-7">
                  <li>시그니처 / 임팩트 조합 시뮬레이션</li>
                  <li>팀 일반 확정권, 팀 고급 확정권 조건 선택</li>
                  <li>위시 카드 등록 및 별 표시</li>
                  <li>일반 자동조합 및 특별 조합</li>
                </ul>
              </section>

              <hr className="border-slate-300" />

              <section>
                <h3 className="text-lg font-black text-slate-900 mb-3">비공식 안내</h3>
                <p className="leading-7 font-semibold">
                  이 사이트는 게임사와 공식적으로 연계되어 있지 않은 비공식 팬 제작 도구입니다.
                  표시되는 시뮬레이션 결과는 참고용이며, 실제 게임 데이터나 업데이트에 따라 달라질 수 있습니다.
                </p>
              </section>
            </div>
          )}

          {infoTab === 'guide' && (
            <div className="space-y-5">
              <section>
                <h2 className="text-xl font-black text-slate-900 mb-3">사용 가이드</h2>
                <ol className="list-decimal pl-5 space-y-2 font-semibold leading-7">
                  <li>상단에서 시그니처 또는 임팩트 조합을 선택합니다.</li>
                  <li>팀 일반 확정권 또는 팀 고급 확정권 조건을 선택합니다.</li>
                  <li>조합 실행 버튼을 눌러 카드 5장을 확인합니다.</li>
                  <li>셔플 시작 후 뒷면 카드 1장을 선택해 최종 결과를 확인합니다.</li>
                  <li>위시 탭에서 원하는 카드를 등록하면 조합 결과에 별 표시가 나타납니다.</li>
                </ol>
              </section>

              <hr className="border-slate-300" />

              <section>
                <h3 className="text-lg font-black text-slate-900 mb-3">자동조합 안내</h3>
                <p className="leading-7 font-semibold">
                  일반 자동조합은 설정한 횟수만큼 빠르게 조합을 실행하고, 1회당 카드 1장을 획득한 것으로 결과를 표시합니다.
                  특별 조합은 특정 카드, 위시 카드, 조합 전용카드가 장판에 등장할 때까지 빠르게 시뮬레이션합니다.
                </p>
              </section>
            </div>
          )}

          {infoTab === 'updates' && (
            <div className="space-y-5">
              <section>
                <h2 className="text-xl font-black text-slate-900 mb-3">업데이트 내역</h2>
                <ul className="space-y-3 font-semibold leading-7">
                  <li>
                    <span className="font-black text-pink-600">v0.5</span> - 커스텀 조합 기능 추가
                  </li>
                  <li>
                    <span className="font-black text-pink-600">v0.4</span> - 모바일 버튼 배치 개선, 확정권 사용 초기화 추가
                  </li>
                  <li>
                    <span className="font-black text-pink-600">v0.3</span> - 위시 등록/관리 기능 추가, 자동조합 기능 정리
                  </li>
                  <li>
                    <span className="font-black text-pink-600">v0.2</span> - 임팩트 조합 탭 추가, 팀 로고 및 모바일 대응
                  </li>
                  <li>
                    <span className="font-black text-pink-600">v0.1</span> - 시그니처 조합 시뮬레이션 기본 기능 구현
                  </li>
                </ul>
              </section>
            </div>
          )}

          {infoTab === 'contact' && (
            <div className="space-y-5">
              <section>
                <h2 className="text-xl font-black text-slate-900 mb-3">문의</h2>
                <p className="leading-7 font-semibold">
                  오류 제보, 카드 DB 수정 요청, 기능 건의는 아래 이메일로 보내주세요.
                </p>

                <a
                  href="mailto:zappa961213@gmail.com"
                  className="mt-4 inline-flex rounded-2xl bg-slate-900 px-5 py-3 font-black text-white hover:bg-pink-600 transition-colors"
                >
                  zappa961213@gmail.com
                </a>
              </section>

              <hr className="border-slate-300" />

              <section>
                <h3 className="text-lg font-black text-slate-900 mb-3">비공식 안내</h3>
                <p className="leading-7 font-semibold">
                  본 사이트는 비공식 팬메이드 시뮬레이터이며, 공식 게임사 또는 구단과 직접적인 관련이 없습니다.
                </p>
              </section>
            </div>
          )}

          <div className="mt-8 text-center text-sm font-bold text-slate-500">
            개발자 주댕
          </div>
        </section>

      </main>
    </div>
  );
}
