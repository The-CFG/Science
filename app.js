// --- 전역 변수 ---
let currentFilename = '';
const BASE_MODAL_Z_INDEX = 70;
const TOP_MODAL_Z_INDEX = 80;
let zIndexCounter = BASE_MODAL_Z_INDEX;

// --- 드래그 관련 변수 ---
let draggedModal = null;
let initialMouseX, initialMouseY;
let initialModalX, initialModalY;

// --- 계산기 관련 상태는 Calculator 모듈 내부에 캡슐화 (하단 참고) ---

// --- 달력 관련 변수 ---
let currentDate = new Date();

/* --- 워드프로세서 모달의 JavaScript 부분 --- */
function formatDoc(command, value = null) {
    document.execCommand(command, false, value);
    document.getElementById('editor').focus();
}

function showSaveDialog() {
    const filenameInput = document.getElementById('filenameInput');
    filenameInput.value = currentFilename || '문서.html';
    document.getElementById('saveUnifiedDialog').classList.remove('hidden');
    document.getElementById('filenameInputSection').classList.remove('hidden');
    document.getElementById('txtWarningSection').classList.add('hidden');
    filenameInput.focus();
    filenameInput.select();
}

function hideSaveUnifiedDialog() {
    document.getElementById('saveUnifiedDialog').classList.add('hidden');
    document.getElementById('dontShowTxtWarningCheckbox').checked = false;
    currentFilename = '';
}

function handleSaveChoice(type) {
    let filename = document.getElementById('filenameInput').value.trim();
    if (!filename) {
        showMessageBox('파일 이름을 입력해주세요.');
        return;
    }
    currentFilename = filename;

    if (type === 'html') {
        saveFile(currentFilename, 'html');
        hideSaveUnifiedDialog();
    } else if (type === 'txt') {
        const dontShowAgain = localStorage.getItem('dontShowTxtWarning');
        if (dontShowAgain === 'true') {
            saveFile(currentFilename, 'txt');
            hideSaveUnifiedDialog();
        } else {
            document.getElementById('filenameInputSection').classList.add('hidden');
            document.getElementById('txtWarningSection').classList.remove('hidden');
        }
    }
}

function confirmTxtSave() {
    const dontShowCheckbox = document.getElementById('dontShowTxtWarningCheckbox');
    if (dontShowCheckbox.checked) {
        localStorage.setItem('dontShowTxtWarning', 'true');
    }
    saveFile(currentFilename, 'txt');
    hideSaveUnifiedDialog();
}

function saveFile(filename, type) {
    let content;
    let mimeType;
    let finalFilename = filename;

    if (type === 'html') {
        content = document.getElementById('editor').innerHTML;
        mimeType = 'text/html;charset=utf-8';
        if (!finalFilename.toLowerCase().endsWith('.html')) {
            finalFilename += '.html';
        }
    } else if (type === 'txt') {
        content = document.getElementById('editor').innerText;
        mimeType = 'text/plain;charset=utf-8';
        if (!finalFilename.toLowerCase().endsWith('.txt')) {
            finalFilename += '.txt';
        }
    } else {
        showMessageBox('유효하지 않은 저장 형식입니다.');
        return;
    }

    if (content) {
        const blob = new Blob([content], { type: mimeType });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = finalFilename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(link.href);
        showMessageBox(`문서가 "${finalFilename}"으로 저장되었습니다.`);
    }
}

function openTextFile(event) {
    const file = event.target.files[0];
    if (file) {
        const fileNameLower = file.name.toLowerCase();
        const isHtml = fileNameLower.endsWith('.html');
        const isTxt = fileNameLower.endsWith('.txt');

        if (!isTxt && !isHtml) {
            showMessageBox('텍스트(.txt) 또는 HTML(.html) 파일만 열 수 있습니다.');
            event.target.value = '';
            return;
        }

        const reader = new FileReader();
        reader.onload = function(e) {
            if (isHtml) {
                document.getElementById('editor').innerHTML = e.target.result;
                showMessageBox('HTML 파일이 로드되었습니다.');
            } else {
                document.getElementById('editor').innerText = e.target.result;
                showMessageBox('텍스트 파일이 로드되었습니다.');
            }
        };
        reader.onerror = function() {
            showMessageBox('파일을 읽는 중 오류가 발생했습니다.');
        };
        reader.readAsText(file);
    }
    event.target.value = '';
}

function showMessageBox(message) {
    const messageBox = document.getElementById('messageBox');
    const messageText = document.getElementById('messageText');
    messageText.innerText = message;
    messageBox.classList.remove('hidden');
}

function hideMessageBox() {
    document.getElementById('messageBox').classList.add('hidden');
}
/* --- /워드프로세서 모달의 JavaScript 부분 --- */


document.addEventListener('DOMContentLoaded', () => {
    // --- 공통 DOM 요소 참조 ---
    const allModals = document.querySelectorAll('.modal-window'); // 일반 앱 모달
    const allGameModals = document.querySelectorAll('.modal'); // 게임 앱 모달
    const allDraggableModals = document.querySelectorAll('.modal-window, .modal'); // 모든 드래그 가능한 모달
    const sidebarEl = document.getElementById('sidebar');

    /* --- 계산기 모달의 DOM 요소 참조 ---
       열기/닫기 버튼은 더 이상 여기서 개별 변수로 선언하지 않음.
       하단의 APP_WINDOWS 설정 배열 + registerAppWindows()가
       id 문자열 하나로 열기/닫기를 한 번에 연결한다. (오타로 인한 버튼 ID 불일치 버그 방지) */
    const calculatorModal = document.getElementById('calculatorModal');
    const display = document.getElementById('display');
    const basicButtons = document.getElementById('basicButtons');
    const advancedButtons = document.getElementById('advancedButtons');
    const toggleCalculatorModeBtn = document.getElementById('toggleCalculatorModeBtn');

    /* --- 달력 모달의 DOM 요소 참조 --- */
    const calendarModal = document.getElementById('calendarModal');
    const currentMonthYear = document.getElementById('currentMonthYear');
    const prevMonthBtn = document.getElementById('prevMonthBtn');
    const nextMonthBtn = document.getElementById('nextMonthBtn');
    const calendarDaysContainer = document.getElementById('calendarDaysContainer');

    /* --- 워드프로세서 모달의 DOM 요소 참조 --- */
    const wordProcessorModal = document.getElementById('wordProcessorModal');

    /* --- 스프레드시트 모달의 DOM 요소 참조 --- */
    const spreadsheetModal = document.getElementById('spreadsheetModal');
    const spreadsheetRoot = document.getElementById('spreadsheet-root');
    let isSpreadsheetAppMounted = false;

    /* --- 캐주얼 게임즈 모달의 DOM 요소 참조 --- */
    const casualGamesModal = document.getElementById('casualGamesModal');
    const launchMinesweeperButton = document.getElementById('launchMinesweeperButton');
    const launchKlondikeButton = document.getElementById('launchKlondikeButton');

    // 지뢰찾기 게임 관련 요소
    let ms_board=[],ms_rows=10,ms_cols=10,ms_bombs=15,ms_flagsPlaced=0,ms_revealedCells=0,ms_gameOver=false,ms_timerInterval=null,ms_seconds=0;
    const minesweeperElements = {
        minesweeperModal: document.getElementById('minesweeper-modal'),
        closeMinesweeperButton: document.getElementById('close-minesweeper-button'),
        minesweeperBoard: document.getElementById('minesweeper-board'),
        flagsRemainingSpan: document.getElementById('flags-remaining'),
        minesweeperTimerSpan: document.getElementById('minesweeper-timer'),
        bombsTotalSpan: document.getElementById('bombs-total'),
        minesweeperResetButton: document.getElementById('minesweeper-reset-button'),
        minesweeperSettingsButton: document.getElementById('minesweeper-settings-button'),
        minesweeperSettingsModal: document.getElementById('minesweeper-settings-modal'),
        minesweeperRowsInput: document.getElementById('minesweeper-rows-input'),
        minesweeperColsInput: document.getElementById('minesweeper-cols-input'),
        minesweeperBombsInput: document.getElementById('minesweeper-bombs-input'),
        minesweeperApplySettingsButton: document.getElementById('minesweeper-apply-settings-button'),
        minesweeperCancelSettingsButton: document.getElementById('minesweeper-cancel-settings-button'),
        minesweeperGameOverMessage: document.getElementById('minesweeper-gameOver-message'),
        minesweeperGameOverTitle: document.getElementById('minesweeper-gameOver-title'),
        minesweeperGameOverText: document.getElementById('minesweeper-gameOver-text'),
        minesweeperPlayAgainButton: document.getElementById('minesweeper-play-again-button'),
        closeMinesweeperSettingsButton: document.getElementById('close-minesweeper-settings-button'),
        closeMinesweeperGameOverButton: document.getElementById('close-minesweeper-gameOver-button')
    };

    // 클론다이크 게임 관련 요소
    const SUITS=['H','D','C','S'],RANKS=['A','2','3','4','5','6','7','8','9','T','J','Q','K'];
    let kl_deck=[],kl_stock=[],kl_waste=[],kl_foundations={'H':[],'D':[],'C':[],'S':[]},kl_tableau=Array(7).fill(null).map(()=>[]);
    let kl_draggedCardsModel=[],kl_sourcePileModel=null,kl_sourcePileIndex=-1;
    let kl_timerInterval=null,kl_seconds=0;
    const klondikeElements = {
        klondikeModal: document.getElementById('klondike-modal'),
        closeKlondikeButton: document.getElementById('close-klondike-button'),
        klondikeStock: document.getElementById('klondike-stock'),
        klondikeWaste: document.getElementById('klondike-waste'),
        klondikeFoundations:{H:document.getElementById('klondike-foundation-H'),D:document.getElementById('klondike-foundation-D'),
                         C:document.getElementById('klondike-foundation-C'),S:document.getElementById('klondike-foundation-S')},
        klondikeTableau:Array(7).fill(null).map((_,i)=>document.getElementById(`klondike-tableau-${i}`)),
        klondikeResetButton: document.getElementById('klondike-reset-button'),
        klondikeTimerSpan: document.getElementById('klondike-timer'),
        klondikeWinMessage: document.getElementById('klondike-win-message'),
        klondikeWinTitle: document.getElementById('klondike-win-title'),
        klondikeWinText: document.getElementById('klondike-win-text'),
        klondikePlayAgainButton: document.getElementById('klondike-play-again-button'),
        closeKlondikeWinButton: document.getElementById('close-klondike-win-button')
    };

    // Klondike Functions (Moved to the top of Klondike section to ensure they are defined before use)
    const createDeck=()=>{let d=[];for(const s of SUITS)for(const r of RANKS)d.push({suit:s,rank:r,faceUp:false});return d;};
    const shuffleDeck=d=>{for(let i=d.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[d[i],d[j]]=[d[j],d[i]];}return d;};
    const getCardValue=r=>RANKS.indexOf(r);
    const isRedSuit=s=>s==='H'||s==='D';
    const isOppositeColor=(c1,c2)=>isRedSuit(c1.suit)!==isRedSuit(c2.suit);
    const getSuitSymbol=s=>{switch(s){case 'H':return'♥';case 'D':return'♦';case 'C':return'♣';case 'S':return'♠';default:return'';}};


    /* --- 도움말 모달의 DOM 요소 참조 --- */
    const helpModal = document.getElementById('helpModal');

    // --- 공통 모달 함수 ---
    const bringModalToFront = (modalElement) => {
        allDraggableModals.forEach(modal => {
            modal.style.zIndex = BASE_MODAL_Z_INDEX;
        });
        modalElement.style.zIndex = TOP_MODAL_Z_INDEX + (++zIndexCounter % 10);
    };

    // 기존 openModal 함수 (modal-window용)
    const openAppModal = (modalElement) => {
        // 이미 위치가 설정되지 않았다면 랜덤 위치 설정
        if (!modalElement.dataset.initialPositioned) {
            modalElement.style.visibility = 'visible';
            modalElement.style.opacity = '0';
            // 데스크탑(768px 이상)에서는 좌측 사이드바 영역과 겹치지 않도록 보정
            const isDesktop = window.innerWidth >= 768;
            const sidebarWidth = isDesktop && sidebarEl ? sidebarEl.offsetWidth : 0;
            const availableWidth = Math.max(window.innerWidth - modalElement.offsetWidth - sidebarWidth - 100, 0);
            const randomX = sidebarWidth + Math.random() * availableWidth + 50;
            const randomY = Math.random() * (window.innerHeight - modalElement.offsetHeight - 100) + 50;
            modalElement.style.left = `${randomX}px`;
            modalElement.style.top = `${randomY}px`;
            modalElement.dataset.initialPositioned = 'true';
            modalElement.style.opacity = '1';
        }
        modalElement.style.visibility = 'visible';
        modalElement.classList.add('active');
        bringModalToFront(modalElement);
    };

    // 기존 closeModal 함수 (modal-window용)
    const closeAppModal = (modalElement) => {
        modalElement.classList.remove('active');
        setTimeout(() => {
            modalElement.style.visibility = 'hidden';
        }, 300);
    };

    // 사이드바 아이콘 아래 "실행 중" 표시 점 토글
    const setDockActive = (openBtn, isActive) => {
        if (!openBtn) return;
        openBtn.classList.toggle('app-active', isActive);
    };

    // 게임 모달용 show/hide 함수 (Casual Games 원본)
    const showGameModal = el => {
        el.classList.add('show');
        el.style.position = 'fixed'; /* 게임 모달은 전체 화면을 덮기 위해 fixed로 유지 */
        el.style.left = '0';
        el.style.top = '0';
        el.style.width = '100%';
        el.style.height = '100%';
        bringModalToFront(el);
    };
    const hideGameModal = el => {
        el.classList.remove('show');
        // 게임 모달이 닫힐 때는 position을 다시 absolute로 바꾸지 않음 (fixed 상태 유지)
    };


    // --- 드래그 기능 구현 ---
    const startDrag = (e) => {
        if (e.type === 'touchstart') {
            e.clientX = e.touches[0].clientX;
            e.clientY = e.touches[0].clientY;
        }
        draggedModal = e.target.closest('.modal-window'); // .modal-window 클래스를 가진 요소만 드래그 가능
        if (!draggedModal) return;

        // 닫기 버튼을 클릭했을 때는 드래그하지 않음
        if (e.target.id && e.target.id.startsWith('close')) {
            draggedModal = null;
            return;
        }

        // 현재 드래그하는 모달을 맨 앞으로 가져옴
        bringModalToFront(draggedModal);

        // 초기 마우스 위치와 모달의 현재 위치를 기록
        initialMouseX = e.clientX;
        initialMouseY = e.clientY;
        initialModalX = draggedModal.offsetLeft;
        initialModalY = draggedModal.offsetTop;

        // 드래그 중에는 커서를 'grabbing'으로 변경
        draggedModal.querySelector('.modal-header').style.cursor = 'grabbing';

        // document 전체에 move 및 up 이벤트 리스너를 추가
        document.addEventListener('mousemove', drag);
        document.addEventListener('mouseup', stopDrag);
        document.addEventListener('touchmove', drag, { passive: false });
        document.addEventListener('touchend', stopDrag);
    };

    const drag = (e) => {
        if (!draggedModal) return;

        // 모바일 터치 이벤트 처리
        if (e.type === 'touchmove') {
            if (e.touches.length === 0) return;
            e.preventDefault(); // 스크롤 방지
            e.clientX = e.touches[0].clientX;
            e.clientY = e.touches[0].clientY;
        }

        // 마우스 이동량 계산
        const dx = e.clientX - initialMouseX;
        const dy = e.clientY - initialMouseY;

        // 모달의 새로운 위치 계산
        let newLeft = initialModalX + dx;
        let newTop = initialModalY + dy;

        // 모달이 화면 밖으로 나가지 않도록 제한 (옵션)
        const viewportWidth = window.innerWidth;
        const viewportHeight = window.innerHeight;
        const modalWidth = draggedModal.offsetWidth;
        const modalHeight = draggedModal.offsetHeight;

        // 최소/최대 좌표 설정 (예: 10px 여백)
        const minLeft = 0;
        const maxLeft = viewportWidth - modalWidth;
        const minTop = 0;
        const maxTop = viewportHeight - modalHeight;

        newLeft = Math.max(minLeft, Math.min(newLeft, maxLeft));
        newTop = Math.max(minTop, Math.min(newTop, maxTop));

        // 모달에 새 위치 적용
        draggedModal.style.left = `${newLeft}px`;
        draggedModal.style.top = `${newTop}px`;
    };

    const stopDrag = () => {
        if (!draggedModal) return;

        // 드래그가 끝났으므로 커서를 기본 'grab'으로 되돌림
        draggedModal.querySelector('.modal-header').style.cursor = 'grab';

        // 전역 드래그 상태를 초기화
        draggedModal = null;

        // document에서 move 및 up 이벤트 리스너를 제거
        document.removeEventListener('mousemove', drag);
        document.removeEventListener('mouseup', stopDrag);
        document.removeEventListener('touchmove', drag);
        document.removeEventListener('touchend', stopDrag);
    };

    // 모든 모달의 헤더에 드래그 시작 이벤트 리스너 추가
    allDraggableModals.forEach(modal => {
        const header = modal.querySelector('.modal-header');
        if (header) { // 헤더가 없는 모달도 있을 수 있으므로 확인
            header.addEventListener('mousedown', startDrag);
            header.addEventListener('touchstart', startDrag, { passive: false });
        }
    });

    /* --- 계산기 모달의 JavaScript 부분 ---
       버그 수정 내역:
       1) 기존에는 계산기 모달을 열 때마다 calculatorModal에 click 리스너를 새로
          추가해서, 여러 번 열고 닫으면 버튼 클릭 한 번에 같은 동작이 여러 번
          실행되는(숫자 중복 입력, 같은 식이 여러 번 계산되는) 문제가 있었음.
          → 리스너는 모듈이 로드될 때 단 한 번만 등록.
       2) 제곱근/로그/삼각함수/π처럼 피연산자가 1개거나 0개인 연산을
          '두 번째 숫자를 입력하고 =를 눌러야' 계산되는 어색한 흐름이었음.
          → 버튼을 누르는 즉시 계산되도록 수정.
       3) 수식 문자열을 eval()로 실행하던 부분을 안전한 switch 함수로 교체.
       4) 소수점(.)을 여러 번 누르면 "3.4.5" 같은 값이 만들어지던 문제 수정. */
    const Calculator = (() => {
        let currentInput = '0';
        let operator = null;
        let prevInput = null;
        let isAdvancedMode = false;

        const BINARY_OPS = ['+', '-', '*', '/', '^'];
        const UNARY_OPS = ['sqrt', 'log', 'sin', 'cos', 'tan'];

        const refreshDisplay = () => { display.value = currentInput; };

        const reset = () => {
            currentInput = '0';
            operator = null;
            prevInput = null;
            refreshDisplay();
        };

        const computeBinary = (a, op, b) => {
            switch (op) {
                case '+': return a + b;
                case '-': return a - b;
                case '*': return a * b;
                case '/': return b === 0 ? NaN : a / b;
                case '^': return Math.pow(a, b);
                default: return NaN;
            }
        };

        const computeUnary = (op, a) => {
            switch (op) {
                case 'sqrt': return Math.sqrt(a);
                case 'log': return Math.log10(a);
                case 'sin': return Math.sin(a * Math.PI / 180);
                case 'cos': return Math.cos(a * Math.PI / 180);
                case 'tan': return Math.tan(a * Math.PI / 180);
                default: return NaN;
            }
        };

        const toDisplayString = (num) => (Number.isFinite(num) ? num.toString() : 'Error');

        const handleButton = (value, mode) => {
            const modeMatches = (mode === 'basic' && !isAdvancedMode) || (mode === 'advanced' && isAdvancedMode);
            if (!modeMatches) return;

            if (value === 'C') { reset(); return; }

            if (value === 'pi') {
                currentInput = Math.PI.toString();
                operator = null;
                prevInput = null;
                refreshDisplay();
                return;
            }

            if (UNARY_OPS.includes(value)) {
                currentInput = toDisplayString(computeUnary(value, parseFloat(currentInput)));
                operator = null;
                prevInput = null;
                refreshDisplay();
                return;
            }

            if (value === '=') {
                if (operator && prevInput !== null) {
                    currentInput = toDisplayString(computeBinary(parseFloat(prevInput), operator, parseFloat(currentInput)));
                    operator = null;
                    prevInput = null;
                }
                refreshDisplay();
                return;
            }

            if (BINARY_OPS.includes(value)) {
                prevInput = prevInput === null
                    ? currentInput
                    : toDisplayString(computeBinary(parseFloat(prevInput), operator, parseFloat(currentInput)));
                operator = value;
                currentInput = '0';
                refreshDisplay();
                return;
            }

            // 숫자 / 소수점 입력
            if (value === '.') {
                if (!currentInput.includes('.')) currentInput += '.';
            } else if (currentInput === '0') {
                currentInput = value;
            } else {
                currentInput += value;
            }
            refreshDisplay();
        };

        const resetUI = () => {
            isAdvancedMode = false;
            basicButtons.classList.remove('hidden');
            advancedButtons.classList.add('hidden');
            toggleCalculatorModeBtn.textContent = '고급 계산기';
            reset();
        };

        const toggleMode = () => {
            isAdvancedMode = !isAdvancedMode;
            basicButtons.classList.toggle('hidden', isAdvancedMode);
            advancedButtons.classList.toggle('hidden', !isAdvancedMode);
            toggleCalculatorModeBtn.textContent = isAdvancedMode ? '기본 계산기' : '고급 계산기';
            reset();
        };

        calculatorModal.addEventListener('click', (e) => {
            const btn = e.target.closest('.calc-button');
            if (!btn) return;
            handleButton(btn.dataset.value, btn.dataset.mode);
        });
        toggleCalculatorModeBtn.addEventListener('click', toggleMode);

        return { resetUI };
    })();
    /* --- /계산기 모달의 JavaScript 부분 --- */

    /* --- 달력 모달의 JavaScript 부분 --- */
    const renderCalendar = () => {
        calendarDaysContainer.innerHTML = '';
        const year = currentDate.getFullYear();
        const month = currentDate.getMonth();

        const monthNames = ["1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월"];
        currentMonthYear.textContent = `${year}년 ${monthNames[month]}`;

        const firstDayOfMonth = new Date(year, month, 1);
        const lastDayOfMonth = new Date(year, month + 1, 0);
        const numDaysInMonth = lastDayOfMonth.getDate();
        const firstDayOfWeek = firstDayOfMonth.getDay();
        const today = new Date();
        const isCurrentMonthToday = today.getFullYear() === year && today.getMonth() === month;

        const prevMonthLastDay = new Date(year, month, 0).getDate();
        for (let i = firstDayOfWeek - 1; i >= 0; i--) {
            const dayElement = document.createElement('div');
            dayElement.classList.add('calendar-day', 'other-month');
            dayElement.textContent = prevMonthLastDay - i;
            calendarDaysContainer.appendChild(dayElement);
        }

        for (let i = 1; i <= numDaysInMonth; i++) {
            const dayElement = document.createElement('div');
            dayElement.classList.add('calendar-day', 'current-month');
            dayElement.textContent = i;
            const dayOfWeek = new Date(year, month, i).getDay();
            if (dayOfWeek === 0) { dayElement.classList.add('weekend', 'text-red-500'); }
            else if (dayOfWeek === 6) { dayElement.classList.add('weekend', 'text-blue-500'); }
            if (isCurrentMonthToday && i === today.getDate()) { dayElement.classList.add('today'); }
            calendarDaysContainer.appendChild(dayElement);
        }

        const totalCellsAfterCurrentMonth = calendarDaysContainer.children.length;
        const maxCells = 35;
        const remainingCells = maxCells - totalCellsAfterCurrentMonth;

        if (remainingCells > 0) {
            for (let i = 1; i <= remainingCells; i++) {
                const dayElement = document.createElement('div');
                dayElement.classList.add('calendar-day', 'other-month');
                dayElement.textContent = i;
                calendarDaysContainer.appendChild(dayElement);
            }
        }
    };

    /* --- 스프레드시트 모달의 JavaScript 부분 (React 코드) --- */
    const FunctionHelpModal = ({ isOpen, onClose }) => {
      if (!isOpen) return null;

      return React.createElement(
        'div',
        { className: 'fixed inset-0 bg-gray-600 bg-opacity-50 flex items-center justify-center p-4 z-[110]' },
        React.createElement(
          'div',
          { className: 'bg-white rounded-lg shadow-xl p-6 max-w-md w-full' },
          React.createElement(
            'h2',
            { className: 'text-2xl font-bold text-gray-800 mb-4' },
            '함수 도움말'
          ),
          React.createElement(
            'p',
            { className: 'text-gray-700 mb-4' },
            '다음 함수를 사용할 수 있습니다:'
          ),
          React.createElement(
            'ul',
            { className: 'list-disc list-inside text-gray-700 mb-6 space-y-2' },
            React.createElement(
              'li',
              null,
              React.createElement('span', { className: 'font-semibold' }, 'SUM(범위)'),
              ': 지정된 범위 내의 모든 숫자를 더합니다.',
              React.createElement('br', null),
              '예시: ',
              React.createElement('code', { className: 'bg-gray-100 px-1 py-0.5 rounded text-sm' }, '=SUM(A1:B5)')
            ),
            React.createElement(
              'li',
              null,
              React.createElement('span', { className: 'font-semibold' }, 'MINUS(값1, 값2)'),
              ': 값1에서 값2를 뺍니다.',
              React.createElement('br', null),
              '예시: ',
              React.createElement('code', { className: 'bg-gray-100 px-1 py-0.5 rounded text-sm' }, '=MINUS(A1,B1)'),
              ' 또는 ',
              React.createElement('code', { className: 'bg-gray-100 px-1 py-0.5 rounded text-sm' }, '=MINUS(100,20)')
            ),
            React.createElement(
              'li',
              null,
              React.createElement('span', { className: 'font-semibold' }, 'TIME(값1, 값2)'),
              ': 값1에 값2를 곱합니다.',
              React.createElement('br', null),
              '예시: ',
              React.createElement('code', { className: 'bg-gray-100 px-1 py-0.5 rounded text-sm' }, '=TIME(A1,B1)'),
              ' 또는 ',
              React.createElement('code', { className: 'bg-gray-100 px-1 py-0.5 rounded text-sm' }, '=TIME(10,5)')
            ),
            React.createElement(
              'li',
              null,
              React.createElement('span', { className: 'font-semibold' }, 'DIVISION(값1, 값2)'),
              ': 값1을 값2로 나눕니다.',
              React.createElement('br', null),
              '예시: ',
              React.createElement('code', { className: 'bg-gray-100 px-1 py-0.5 rounded text-sm' }, '=DIVISION(A1,B1)'),
              ' 또는 ',
              React.createElement('code', { className: 'bg-gray-100 px-1 py-0.5 rounded text-sm' }, '=DIVISION(100,10)')
            )
          ),
          React.createElement(
            'div',
            { className: 'flex justify-end' },
            React.createElement(
              'button',
              {
                onClick: onClose,
                className: 'px-6 py-2 bg-gray-300 text-gray-800 rounded-md hover:bg-gray-400 focus:outline-none focus:ring-2 focus:ring-gray-300 transition duration-150 ease-in-out'
              },
              '닫기'
            )
          )
        )
      );
    };

    const SpreadsheetApp = () => {
      const { useState, useEffect, useRef } = React;

      const [focusedCell, setFocusedCell] = useState(null);
      const [isSpreadsheetHelpModalOpen, setIsSpreadsheetHelpModalOpen] = useState(false);

      const [gridData, setGridData] = useState(() => {
        const initialRawData = [];
        for (let r = 0; r < 10; r++) {
          initialRawData.push(Array(5).fill(''));
        }
        return initialRawData.map(row => row.map(rawVal => ({ raw: rawVal, display: rawVal })));
      });

      const numRows = gridData.length;
      const numCols = gridData.length > 0 ? gridData[0].length : 0;

      const parseCellReference = (ref) => {
        const colCharMatch = ref.match(/[A-Z]/i);
        const rowNumMatch = ref.match(/\d+/);

        if (!colCharMatch || !rowNumMatch) {
          throw new Error('Invalid cell reference format: ' + ref);
        }

        const col = colCharMatch[0].toUpperCase().charCodeAt(0) - 65;
        const row = parseInt(rowNumMatch[0], 10) - 1;

        if (col < 0 || row < 0) {
          throw new Error('Cell reference out of bounds (negative indices): ' + ref);
        }
        return { col, row };
      };

      const parseRange = (rangeStr) => {
        const parts = rangeStr.split(':');
        const startRef = parseCellReference(parts[0]);
        const endRef = parts.length > 1 ? parseCellReference(parts[1]) : startRef;

        return {
          startCol: Math.min(startRef.col, endRef.col),
          startRow: Math.min(startRef.row, endRef.row),
          endCol: Math.max(startRef.col, endRef.col),
          endRow: Math.max(startRef.row, endRef.row)
        };
      };

      const getNumericValue = (arg, currentGridState) => {
        try {
          const { row, col } = parseCellReference(arg);
          if (row >= 0 && row < numRows && col >= 0 && col < numCols) {
            const { display: referencedDisplayValue } = evaluateCell(row, col, currentGridState);
            const val = parseFloat(referencedDisplayValue);
            return isNaN(val) ? 0 : val;
          }
          return 0;
        } catch (e) {
          const num = parseFloat(arg);
          return isNaN(num) ? 0 : num;
        }
      };

      const evaluateCell = (rowIndex, colIndex, currentGridState) => {
        if (rowIndex >= numRows || colIndex >= numCols || rowIndex < 0 || colIndex < 0) {
          return { display: '#REF!' };
        }
        if (!currentGridState[rowIndex] || !currentGridState[rowIndex][colIndex]) {
          return { display: '#REF!' };
        }

        const rawValue = currentGridState[rowIndex][colIndex].raw;

        if (typeof rawValue !== 'string' || !rawValue.startsWith('=')) {
          return { display: rawValue };
        }

        try {
          const formula = rawValue.substring(1).trim();
          const upperFormula = formula.toUpperCase();

          if (upperFormula.startsWith('SUM(') && upperFormula.endsWith(')')) {
            const rangeStr = formula.substring(4, formula.length - 1);
            const { startCol, startRow, endCol, endRow } = parseRange(rangeStr);

            let sum = 0;
            for (let r = startRow; r <= endRow; r++) {
              for (let c = startCol; c <= endCol; c++) {
                if (r < numRows && c < numCols && currentGridState[r] && currentGridState[r][c]) {
                  const { display: referencedDisplayValue } = evaluateCell(r, c, currentGridState);
                  const val = parseFloat(referencedDisplayValue);
                  if (!isNaN(val)) {
                    sum += val;
                  }
                } else {
                  console.warn(`Cell reference out of bounds or invalid in SUM formula: ${String.fromCharCode(65 + c)}${r + 1}`);
                }
              }
            }
            return { display: sum.toString() };
          }
          else if (upperFormula.startsWith('MINUS(') || upperFormula.startsWith('TIME(') || upperFormula.startsWith('DIVISION(')) {
            const funcNameEndIndex = upperFormula.indexOf('(');
            const funcName = upperFormula.substring(0, funcNameEndIndex);
            const argsStr = formula.substring(funcNameEndIndex + 1, formula.length - 1);
            const args = argsStr.split(',').map(arg => arg.trim());

            if (args.length !== 2) {
              return { display: '#VALUE!' };
            }

            const val1 = getNumericValue(args[0], currentGridState);
            const val2 = getNumericValue(args[1], currentGridState);

            let result;
            switch (funcName) {
              case 'MINUS':
                result = val1 - val2;
                break;
              case 'TIME':
                result = val1 * val2;
                break;
              case 'DIVISION':
                if (val2 === 0) {
                  return { display: '#DIV/0!' };
                }
                result = val1 / val2;
                break;
              default:
                return { display: '#NAME?' };
            }
            return { display: result.toString() };
          }
          else {
            return { display: '#NAME?' };
          }
        } catch (e) {
          console.error('Formula evaluation error:', e);
          return { display: '#ERROR!' };
        }
      };

      const handleCellChange = (rowIndex, colIndex, value) => {
        setGridData(prevGridData => {
          const newGridData = prevGridData.map(row => row.map(cell => ({ ...cell })));
          newGridData[rowIndex][colIndex].raw = value;

          const updatedGridWithDisplays = newGridData.map((row, rIdx) =>
            row.map((cell, cIdx) => {
              const { display } = evaluateCell(rIdx, cIdx, newGridData);
              return { raw: cell.raw, display };
            })
          );
          return updatedGridWithDisplays;
        });
      };

      const addRow = () => {
        setGridData(prevGridData => {
          const newRow = Array(numCols).fill({ raw: '', display: '' });
          const newGrid = [...prevGridData, newRow];
          return newGrid.map((row, rIdx) =>
            row.map((cell, cIdx) => {
              const { display } = evaluateCell(rIdx, cIdx, newGrid);
              return { raw: cell.raw, display };
            })
          );
        });
      };

      const addColumn = () => {
        setGridData(prevGridData => {
          const newGrid = prevGridData.map(row => [...row, { raw: '', display: '' }]);
          return newGrid.map((row, rIdx) =>
            row.map((cell, cIdx) => {
              const { display } = evaluateCell(rIdx, cIdx, newGrid);
              return { raw: cell.raw, display };
            })
          );
        });
      };

      return React.createElement(
        'div',
        { className: 'min-h-full bg-gray-100 p-4 font-inter flex flex-col' },
        React.createElement(
          'div',
          { className: 'container mx-auto bg-white shadow-lg rounded-lg p-6 flex flex-col flex-grow' },
          React.createElement(
            'h1',
            { className: 'text-3xl font-bold text-gray-800 mb-6 text-center' },
            '간단한 스프레드시트'
          ),
          React.createElement(
            'div',
            { className: 'mb-4 flex justify-center space-x-4' },
            React.createElement(
              'button',
              {
                onClick: addRow,
                className: 'px-6 py-3 bg-blue-500 text-white rounded-md hover:bg-blue-600 focus:outline-none focus:ring-2 focus:ring-blue-400 focus:ring-opacity-75 transition duration-150 ease-in-out shadow-md'
              },
              '행 추가'
            ),
            React.createElement(
              'button',
              {
                onClick: addColumn,
                className: 'px-6 py-3 bg-green-500 text-white rounded-md hover:bg-green-600 focus:outline-none focus:ring-2 focus:ring-green-400 focus:ring-opacity-75 transition duration-150 ease-in-out shadow-md'
              },
              '열 추가'
            ),
            React.createElement(
              'button',
              {
                onClick: () => setIsSpreadsheetHelpModalOpen(true),
                className: 'px-6 py-3 bg-purple-500 text-white rounded-md hover:bg-purple-600 focus:outline-none focus:ring-2 focus:ring-purple-400 focus:ring-opacity-75 transition duration-150 ease-in-out shadow-md'
              },
              '함수 도움말'
            )
          ),
          React.createElement(
            'div',
            { className: 'overflow-x-auto overflow-y-auto flex-grow' },
            React.createElement(
              'table',
              { className: 'min-w-full border-collapse table-fixed' },
              React.createElement(
                'thead',
                null,
                React.createElement(
                  'tr',
                  { className: 'bg-gray-200' },
                  React.createElement('th', { className: 'px-3 py-2 border border-gray-300 text-center text-sm font-semibold text-gray-600 rounded-tl-md sticky top-0 left-0 bg-gray-200 z-10' }),
                  Array.from({ length: numCols }).map((_, colIndex) =>
                    React.createElement(
                      'th',
                      {
                        key: `col-header-${colIndex}`,
                        className: `px-4 py-2 border border-gray-300 text-center text-sm font-semibold text-gray-600 sticky top-0 bg-gray-200 z-10 ${
                          colIndex === numCols - 1 ? 'rounded-tr-md' : ''
                        }`
                      },
                      String.fromCharCode(65 + colIndex)
                    )
                  )
                )
              ),
              React.createElement(
                'tbody',
                null,
                Array.from({ length: numRows }).map((_, rowIndex) =>
                  React.createElement(
                    'tr',
                    { key: `row-${rowIndex}`, className: 'hover:bg-gray-50' },
                    React.createElement(
                      'td',
                      { className: 'px-3 py-2 border border-gray-300 text-center text-sm font-semibold text-gray-600 bg-gray-200 sticky left-0 z-10' },
                      rowIndex + 1
                    ),
                    Array.from({ length: numCols }).map((_, colIndex) =>
                      React.createElement(
                        'td',
                        { key: `cell-${rowIndex}-${colIndex}`, className: 'p-0 border border-gray-300' },
                        React.createElement('input', {
                          type: 'text',
                          value:
                            focusedCell && focusedCell.row === rowIndex && focusedCell.col === colIndex
                              ? gridData[rowIndex][colIndex].raw
                              : gridData[rowIndex][colIndex].display,
                          onChange: (e) => handleCellChange(rowIndex, colIndex, e.target.value),
                          onFocus: () => setFocusedCell({ row: rowIndex, col: colIndex }),
                          onBlur: () => setFocusedCell(null),
                          className: 'w-full h-full p-2 outline-none focus:ring-2 focus:ring-blue-300 rounded-md transition duration-150 ease-in-out text-gray-700 text-sm',
                          spellCheck: 'false'
                        })
                      )
                    )
                  )
                )
              )
            )
          )
        ),
        React.createElement(FunctionHelpModal, {
          isOpen: isSpreadsheetHelpModalOpen,
          onClose: () => setIsSpreadsheetHelpModalOpen(false)
        })
      );
    };

    // 스프레드시트 앱 초기화 함수
    const spreadsheetModalInit = () => {
        if (!isSpreadsheetAppMounted) {
            if (typeof ReactDOM === 'undefined' || typeof React === 'undefined') {
                spreadsheetRoot.innerHTML = '<div style="padding:2rem;text-align:center;color:#6b7280"><p style="font-size:1.2rem;margin-bottom:0.5rem">⚠️ 스프레드시트를 로드할 수 없습니다</p><p style="font-size:0.9rem">React 라이브러리를 불러오지 못했습니다.<br>인터넷 연결을 확인하고 페이지를 새로 고침해 주세요.</p></div>';
                return;
            }
            ReactDOM.render(
                React.createElement(SpreadsheetApp),
                spreadsheetRoot
            );
            isSpreadsheetAppMounted = true;
        }
    };
    /* --- /스프레드시트 모달의 JavaScript 부분 (React 코드) --- */

    /* --- 캐주얼 게임즈 모달의 JavaScript 부분 --- */
    // 지뢰찾기 함수들 (minesweeperElements 사용)
    const getMinesweeperCellElement=(r,c)=>minesweeperElements.minesweeperBoard.children[r*ms_cols+c];
    function initMinesweeper(){
        ms_gameOver=false;ms_flagsPlaced=0;ms_revealedCells=0;ms_seconds=0;
        if(ms_timerInterval)clearInterval(ms_timerInterval);ms_timerInterval=null;
        minesweeperElements.minesweeperBoard.innerHTML='';minesweeperElements.minesweeperBoard.style.gridTemplateColumns=`repeat(${ms_cols},1fr)`;
        minesweeperElements.bombsTotalSpan.textContent=ms_bombs;minesweeperElements.flagsRemainingSpan.textContent=ms_bombs-ms_flagsPlaced;minesweeperElements.minesweeperTimerSpan.textContent=ms_seconds;
        ms_board=Array(ms_rows).fill(null).map(()=>Array(ms_cols).fill(null).map(()=>({isBomb:false,isRevealed:false,isFlagged:false,neighborBombs:0})));
        let placed=0;while(placed<ms_bombs){const r=Math.floor(Math.random()*ms_rows),c=Math.floor(Math.random()*ms_cols);if(!ms_board[r][c].isBomb){ms_board[r][c].isBomb=true;placed++;}}
        for(let r=0;r<ms_rows;r++)for(let c=0;c<ms_cols;c++)if(!ms_board[r][c].isBomb){let count=0;for(let dr=-1;dr<=1;dr++)for(let dc=-1;dc<=1;dc++){const nr=r+dr,nc=c+dc;if(nr>=0&&nr<ms_rows&&nc>=0&&nc<ms_cols&&ms_board[nr][nc].isBomb)count++;}ms_board[r][c].neighborBombs=count;}
        for(let r=0;r<ms_rows;r++){for(let c=0;c<ms_cols;c++){const cell=document.createElement('div');cell.classList.add('cell','rounded');cell.dataset.row=r;cell.dataset.col=c;
            let longPressTimer=null,isTouchMoved=false,isLongPressActivated=false;
            cell.addEventListener('mousedown',e=>{if(e.button===0)revealMinesweeperCell(r,c);});
            cell.addEventListener('contextmenu',e=>{e.preventDefault();toggleMinesweeperFlag(r,c);});
            cell.addEventListener('touchstart',e=>{
                e.preventDefault();isTouchMoved=false;isLongPressActivated=false;
                cell.dataset.initialTouchX=e.touches[0].clientX;cell.dataset.initialTouchY=e.touches[0].clientY;
                longPressTimer=setTimeout(()=>{isLongPressActivated=true;if(!ms_gameOver&&!ms_board[r][c].isRevealed)toggleMinesweeperFlag(r,c);longPressTimer=null;},500);
            },{passive:false});
            cell.addEventListener('touchmove',e=>{
                const curX=e.touches[0].clientX,curY=e.touches[0].clientY;
                const initX=parseFloat(cell.dataset.initialTouchX||0),initY=parseFloat(cell.dataset.initialTouchY||0);
                if(Math.abs(curX-initX)>10||Math.abs(curY-initY)>10){isTouchMoved=true;if(longPressTimer){clearTimeout(longPressTimer);longPressTimer=null;}isLongPressActivated=false;}
            },{passive:false});
            cell.addEventListener('touchend',e=>{
                if(longPressTimer){clearTimeout(longPressTimer);longPressTimer=null;}
                if(!isLongPressActivated&&!isTouchMoved){if(!ms_gameOver&&!ms_board[r][c].isRevealed&&!ms_board[r][c].isFlagged)revealMinesweeperCell(r,c);}
                isTouchMoved=false;isLongPressActivated=false;
            });
            cell.addEventListener('touchcancel',()=>{if(longPressTimer)clearTimeout(longPressTimer);longPressTimer=null;isTouchMoved=false;isLongPressActivated=false;});
            minesweeperElements.minesweeperBoard.appendChild(cell);
        }}
        hideGameModal(minesweeperElements.minesweeperGameOverMessage);
    }
    function revealMinesweeperCell(r,c){
        if(ms_gameOver||ms_board[r][c].isRevealed||ms_board[r][c].isFlagged)return;
        if(ms_revealedCells===0&&ms_timerInterval===null)ms_timerInterval=setInterval(()=>{ms_seconds++;minesweeperElements.minesweeperTimerSpan.textContent=ms_seconds;},1000);
        ms_board[r][c].isRevealed=true;ms_revealedCells++;
        const cellEl=getMinesweeperCellElement(r,c);cellEl.classList.add('revealed');
        if(ms_board[r][c].isBomb){cellEl.classList.add('bomb','exploded');endMinesweeperGame(false);return;}
        if(ms_board[r][c].neighborBombs>0){cellEl.textContent=ms_board[r][c].neighborBombs;cellEl.dataset.value=ms_board[r][c].neighborBombs;}
        else{for(let dr=-1;dr<=1;dr++)for(let dc=-1;dc<=1;dc++){const nr=r+dr,nc=c+dc;if(nr>=0&&nr<ms_rows&&nc>=0&&nc<ms_cols&&!ms_board[nr][nc].isRevealed)revealMinesweeperCell(nr,nc);}}
        checkMinesweeperWin();
    }
    function toggleMinesweeperFlag(r,c){
        if(ms_gameOver||ms_board[r][c].isRevealed)return;
        const cellEl=getMinesweeperCellElement(r,c);
        if(ms_board[r][c].isFlagged){ms_board[r][c].isFlagged=false;ms_flagsPlaced--;cellEl.classList.remove('flag');cellEl.innerHTML='';}
        else{if(ms_flagsPlaced<ms_bombs){ms_board[r][c].isFlagged=true;ms_flagsPlaced++;cellEl.classList.add('flag');cellEl.innerHTML='<i class="fas fa-flag"></i>';}}
        minesweeperElements.flagsRemainingSpan.textContent=ms_bombs-ms_flagsPlaced;
    }
    function endMinesweeperGame(win){
        ms_gameOver=true;if(ms_timerInterval){clearInterval(ms_timerInterval);ms_timerInterval=null;}
        for(let r=0;r<ms_rows;r++)for(let c=0;c<ms_cols;c++){const cellEl=getMinesweeperCellElement(r,c);if(ms_board[r][c].isBomb&&!ms_board[r][c].isFlagged){cellEl.classList.add('bomb');cellEl.innerHTML='<i class="fas fa-bomb"></i>';}else if(ms_board[r][c].isFlagged&&!ms_board[r][c].isBomb){cellEl.innerHTML='<i class="fas fa-times" style="color:red;"></i>';}if(!ms_board[r][c].isRevealed)cellEl.classList.add('revealed');cellEl.style.cursor='default';}
        showGameModal(minesweeperElements.minesweeperGameOverMessage); // 게임 오버 모달은 showGameModal 사용
        minesweeperElements.minesweeperGameOverTitle.textContent=win?'축하합니다!':'게임 오버!';
        minesweeperElements.minesweeperGameOverText.textContent=win?`모든 지뢰를 찾았습니다! ${ms_seconds}초가 걸렸습니다.`:'지뢰를 밟았습니다. 다시 시도하세요!';
        minesweeperElements.minesweeperGameOverTitle.style.color=win?'#10B981':'#EF4444';
    }
    function checkMinesweeperWin(){if(ms_revealedCells===(ms_rows*ms_cols-ms_bombs))endMinesweeperGame(true);}
    const launchMinesweeper=()=>{
        hideGameModal(casualGamesModal); // 캐주얼 게임 메인 모달 숨기기
        showGameModal(minesweeperElements.minesweeperModal);initMinesweeper();
    };
    const closeMinesweeper=()=>{
        hideGameModal(minesweeperElements.minesweeperModal);openAppModal(casualGamesModal); // 캐주얼 게임 메인 모달 다시 열기
        if(ms_timerInterval)clearInterval(ms_timerInterval);ms_timerInterval=null;ms_seconds=0;minesweeperElements.minesweeperTimerSpan.textContent=ms_seconds;
    };

    // 클론다이크 함수들 (klondikeElements 사용)
    const renderKlondike=()=>{
        klondikeElements.klondikeStock.innerHTML='';
        if(kl_stock.length>0){const cEl=createCardElement(kl_stock[kl_stock.length-1],true);klondikeElements.klondikeStock.appendChild(cEl);}
        klondikeElements.klondikeWaste.innerHTML='';
        if(kl_waste.length>0){const c=kl_waste[kl_waste.length-1];const cEl=createCardElement(c,false);klondikeElements.klondikeWaste.appendChild(cEl);cEl.style.position='relative';cEl.style.left='0';cEl.style.top='0';addDragListeners(cEl,kl_waste,kl_waste.length-1);}
        for(const s of SUITS){
            klondikeElements.klondikeFoundations[s].innerHTML='';
            if(kl_foundations[s].length>0){const c=kl_foundations[s][kl_foundations[s].length-1];const cEl=createCardElement(c,false);klondikeElements.klondikeFoundations[s].appendChild(cEl);cEl.style.position='relative';cEl.style.left='0';cEl.style.top='0';}
        }
        kl_tableau.forEach((p,pIdx)=>{
            const tEl=klondikeElements.klondikeTableau[pIdx];tEl.innerHTML='';
            p.forEach((c,cIdx)=>{
                const cEl=createCardElement(c,!c.faceUp);
                cEl.style.position='absolute';cEl.style.top=`${cIdx*20}px`;cEl.style.left='0';
                if(c.faceUp)addDragListeners(cEl,p,cIdx);
                tEl.appendChild(cEl);
            });
        });
    };

    const createCardElement=(c,isFD)=>{
        const cEl=document.createElement('div');cEl.classList.add('card','rounded');
        if(isFD)cEl.classList.add('face-down');
        else{cEl.innerHTML=`<div class="rank">${c.rank}</div><div class="suit">${getSuitSymbol(c.suit)}</div>`;cEl.classList.add(isRedSuit(c.suit)?'red':'black');}
        cEl.dataset.suit=c.suit;cEl.dataset.rank=c.rank;return cEl;
    };

    const initKlondike=()=>{
        kl_deck=shuffleDeck(createDeck());kl_stock=[];kl_waste=[];kl_foundations={'H':[],'D':[],'C':[],'S':[]};kl_tableau=Array(7).fill(null).map(()=>[]);
        for(let i=0;i<7;i++)for(let j=0;j<=i;j++){const c=kl_deck.pop();c.faceUp=(j===i);kl_tableau[i].push(c);}
        kl_stock=kl_deck;
        renderKlondike();startKlondikeTimer();hideGameModal(klondikeElements.klondikeWinMessage);
    };
    const startKlondikeTimer=()=>{
        if(kl_timerInterval)clearInterval(kl_timerInterval);kl_seconds=0;klondikeElements.klondikeTimerSpan.textContent=kl_seconds;
        kl_timerInterval=setInterval(()=>{kl_seconds++;klondikeElements.klondikeTimerSpan.textContent=kl_seconds;},1000);
    };
    const stopKlondikeTimer=()=>{if(kl_timerInterval){clearInterval(kl_timerInterval);kl_timerInterval=null;}};

    const addDragListeners=(cardEl,pile,cardIndex)=>{
        let sx,sy,il,it,currentDragDOMEl=null,draggedCardsDOMGroup=[];
        const onMM=e=>{
            if(!currentDragDOMEl)return;e.preventDefault();
            const cx=e.touches?e.touches[0].clientX:e.clientX,cy=e.touches?e.touches[0].clientY:e.clientY;
            const dx=cx-sx,dy=cy-sy;
            draggedCardsDOMGroup.forEach((cEl,idx)=>{
                const oTO=(pile===kl_waste)?0:(cardIndex*20);
                cEl.style.left=`${il+dx}px`;cEl.style.top=`${oTO+dy+(idx*20)}px`;
            });
        };
        const onMU=e=>{
            if(!currentDragDOMEl)return;
            document.removeEventListener('mousemove',onMM);document.removeEventListener('mouseup',onMU);
            document.removeEventListener('touchmove',onMM);document.removeEventListener('touchend',onMU);document.removeEventListener('touchcancel',onMU);
            currentDragDOMEl.classList.remove('dragging');draggedCardsDOMGroup.forEach(c=>c.classList.remove('dragging'));
            const dropTarget=getDropTarget(currentDragDOMEl);
            if(dropTarget){
                attemptMove(kl_draggedCardsModel,kl_sourcePileModel,kl_sourcePileIndex,dropTarget);
            }else{
                draggedCardsDOMGroup.forEach((cEl,idx)=>{
                    const oTO=(kl_sourcePileModel===kl_waste)?0:(kl_sourcePileIndex*20);
                    cEl.style.left='0px';cEl.style.top=`${oTO+(idx*20)}px`;
                    cEl.style.position=(kl_sourcePileModel===kl_waste&&kl_sourcePileModel.length===1&&idx===0)?'relative':'absolute';
                });
            }
            currentDragDOMEl=null;draggedCardsDOMGroup=[];kl_draggedCardsModel=[];kl_sourcePileModel=null;kl_sourcePileIndex=-1;
            renderKlondike();
        };
        cardEl.addEventListener('mousedown',e=>{
            if(e.button!==0||cardEl.classList.contains('face-down'))return;e.preventDefault();
            sx=e.clientX;sy=e.clientY;il=cardEl.offsetLeft;it=cardEl.offsetTop;
            currentDragDOMEl=cardEl;currentDragDOMEl.classList.add('dragging');
            if(pile===kl_waste){
                kl_draggedCardsModel=[pile[cardIndex]];kl_sourcePileModel=kl_waste;kl_sourcePileIndex=cardIndex;draggedCardsDOMGroup=[cardEl];
            }else{
                kl_draggedCardsModel=pile.slice(cardIndex);kl_sourcePileModel=pile;kl_sourcePileIndex=cardIndex;draggedCardsDOMGroup=Array.from(cardEl.parentNode.children).filter((el,idx)=>idx>=cardIndex);
            }
            draggedCardsDOMGroup.forEach((cEl,idx)=>{
                cEl.style.position='absolute';cEl.style.zIndex=500+idx;
                const oTO=(kl_sourcePileModel===kl_waste)?0:(kl_sourcePileIndex*20);
                cEl.style.left=`${il}px`;cEl.style.top=`${oTO+(idx*20)}px`;
            });
            document.addEventListener('mousemove',onMM);document.addEventListener('mouseup',onMU);
        });
        cardEl.addEventListener('touchstart',e=>{
            if(cardEl.classList.contains('face-down'))return;e.preventDefault();
            sx=e.touches[0].clientX;sy=e.touches[0].clientY;il=cardEl.offsetLeft;it=cardEl.offsetTop;
            currentDragDOMEl=cardEl;currentDragDOMEl.classList.add('dragging');
            if(pile===kl_waste){
                kl_draggedCardsModel=[pile[cardIndex]];kl_sourcePileModel=kl_waste;kl_sourcePileIndex=cardIndex;draggedCardsDOMGroup=[cardEl];
            }else{
                kl_draggedCardsModel=pile.slice(cardIndex);kl_sourcePileModel=pile;kl_sourcePileIndex=cardIndex;draggedCardsDOMGroup=Array.from(cardEl.parentNode.children).filter((el,idx)=>idx>=cardIndex);
            }
            draggedCardsDOMGroup.forEach((cEl,idx)=>{
                cEl.style.position='absolute';cEl.style.zIndex=500+idx;
                const oTO=(kl_sourcePileModel===kl_waste)?0:(kl_sourcePileIndex*20);
                cEl.style.left=`${il}px`;cEl.style.top=`${oTO+(idx*20)}px`;
            });
            document.addEventListener('touchmove',onMM,{passive:false});document.addEventListener('touchend',onMU);document.addEventListener('touchcancel',onMU);
        },{passive:false});
    };

    const getDropTarget=draggedEl=>{
        const dR=draggedEl.getBoundingClientRect(),center={x:dR.left+dR.width/2,y:dR.top+dR.height/2};
        for(const s of SUITS){
            const fEl=klondikeElements.klondikeFoundations[s],fR=fEl.getBoundingClientRect();
            if(center.x>fR.left&&center.x<fR.right&&center.y>fR.top&&center.y<fR.bottom)return{type:'foundation',pile:kl_foundations[s],suit:s,el:fEl};
        }
        for(let i=0;i<7;i++){
            const tEl=klondikeElements.klondikeTableau[i],tR=tEl.getBoundingClientRect();
            const lastCardEl=tEl.children[tEl.children.length-1];
            const eB=lastCardEl?lastCardEl.getBoundingClientRect().bottom:tR.bottom;
            if(center.x>tR.left&&center.x<tR.right&&center.y>tR.top&&center.y<eB+20)return{type:'tableau',pile:kl_tableau[i],index:i,el:tEl};
        }
        return null;
    };

    const attemptMove=(draggedModelCards,sourcePileArray,sourcePileStartIndex,target)=>{
        if(draggedModelCards.length===0)return;
        const fC=draggedModelCards[0];let isValid=false;
        if(target.type==='foundation'){
            if(draggedModelCards.length===1){
                const tFC=target.pile.length>0?target.pile[target.pile.length-1]:null;
                if((!tFC&&fC.rank==='A')||(tFC&&fC.suit===tFC.suit&&getCardValue(fC.rank)===getCardValue(tFC.rank)+1))isValid=true;
            }
        }else if(target.type==='tableau'){
            const tTC=target.pile.length>0?target.pile[target.pile.length-1]:null;
            if((!tTC&&fC.rank==='K')||(tTC&&isOppositeColor(fC,tTC)&&getCardValue(fC.rank)===getCardValue(tTC.rank)-1))isValid=true;
        }
        if(isValid){
            const rC=sourcePileArray.splice(sourcePileStartIndex,draggedModelCards.length);
            target.pile.push(...rC);
            if(kl_tableau.includes(sourcePileArray)){if(sourcePileArray.length>0&&!sourcePileArray[sourcePileArray.length-1].faceUp){sourcePileArray[sourcePileArray.length-1].faceUp=true;}}
            checkKlondikeWin();
        }
    };

    // 버그 수정: 이전에는 showKlondikeWinMessage()가 정의만 되어 있고 어디서도 호출되지
    // 않아서, 승리 모달이 떠도 "몇 초 만에 완료했습니다" 안내 문구가 항상 빈 칸이었음.
    function checkKlondikeWin(){
        const allFoundationsComplete=SUITS.every(s=>kl_foundations[s].length===13);
        if(allFoundationsComplete){
            stopKlondikeTimer();
            klondikeElements.klondikeWinText.textContent=`축하합니다! 클론다이크를 ${kl_seconds}초 만에 완료했습니다!`;
            showGameModal(klondikeElements.klondikeWinMessage);
        }
    }

    const launchKlondike=()=>{
        hideGameModal(casualGamesModal); // 캐주얼 게임 메인 모달 숨기기
        showGameModal(klondikeElements.klondikeModal);initKlondike();
    };
    const closeKlondike=()=>{
        hideGameModal(klondikeElements.klondikeModal);openAppModal(casualGamesModal); // 캐주얼 게임 메인 모달 다시 열기
        stopKlondikeTimer();
    };
    /* --- /캐주얼 게임즈 모달의 JavaScript 부분 --- */


    /* --- 각 앱 창(모달) 열기/닫기 연결은 파일 하단의 APP_WINDOWS 설정으로 일괄 처리 --- */

    /* --- 달력 모달의 JavaScript 부분 (월 이동 버튼) --- */
    prevMonthBtn.addEventListener('click', () => { currentDate.setMonth(currentDate.getMonth() - 1); renderCalendar(); });
    nextMonthBtn.addEventListener('click', () => { currentDate.setMonth(currentDate.getMonth() + 1); renderCalendar(); });

    /* --- 캐주얼 게임즈: 게임 실행 버튼 --- */
    launchMinesweeperButton.addEventListener('click', launchMinesweeper);
    launchKlondikeButton.addEventListener('click', launchKlondike);

    // 지뢰찾기 내부 버튼 이벤트 리스너
    minesweeperElements.minesweeperResetButton.addEventListener('click',initMinesweeper);
    minesweeperElements.minesweeperPlayAgainButton.addEventListener('click',()=>{hideGameModal(minesweeperElements.minesweeperGameOverMessage);initMinesweeper();});
    minesweeperElements.minesweeperSettingsButton.addEventListener('click',()=>{
        minesweeperElements.minesweeperRowsInput.value=ms_rows;minesweeperElements.minesweeperColsInput.value=ms_cols;minesweeperElements.minesweeperBombsInput.value=ms_bombs;
        showGameModal(minesweeperElements.minesweeperSettingsModal);
    });
    minesweeperElements.minesweeperApplySettingsButton.addEventListener('click',()=>{
        const r=parseInt(minesweeperElements.minesweeperRowsInput.value),c=parseInt(minesweeperElements.minesweeperColsInput.value),b=parseInt(minesweeperElements.minesweeperBombsInput.value);
        if(isNaN(r)||r<5||r>30||isNaN(c)||c<5||c>30||isNaN(b)||b<1){console.error('Invalid input.');return;}
        if(b>=r*c){console.error('Bombs cannot exceed cells.');return;}
        ms_rows=r;ms_cols=c;ms_bombs=b;
        minesweeperElements.minesweeperBoard.style.gridTemplateColumns=`repeat(${ms_cols},1fr)`;
        hideGameModal(minesweeperElements.minesweeperSettingsModal);initMinesweeper();
    });
    minesweeperElements.minesweeperCancelSettingsButton.addEventListener('click',()=>hideGameModal(minesweeperElements.minesweeperSettingsModal));
    minesweeperElements.closeMinesweeperButton.addEventListener('click',closeMinesweeper); // 게임 모달 닫기
    minesweeperElements.closeMinesweeperSettingsButton.addEventListener('click',()=>hideGameModal(minesweeperElements.minesweeperSettingsModal));
    minesweeperElements.closeMinesweeperGameOverButton.addEventListener('click',()=>hideGameModal(minesweeperElements.minesweeperGameOverMessage));

    // 클론다이크 내부 버튼 이벤트 리스너
    klondikeElements.klondikeStock.addEventListener('click',()=>{
        if(kl_stock.length>0){const c=kl_stock.pop();c.faceUp=true;kl_waste.push(c);}
        else if(kl_waste.length>0){kl_stock=kl_waste.reverse().map(c=>({...c,faceUp:false}));kl_waste=[];}
        renderKlondike();
    });
    klondikeElements.klondikeResetButton.addEventListener('click',initKlondike);
    klondikeElements.klondikePlayAgainButton.addEventListener('click',()=>{hideGameModal(klondikeElements.klondikeWinMessage);initKlondike();});
    klondikeElements.closeKlondikeButton.addEventListener('click',closeKlondike); // 게임 모달 닫기
    klondikeElements.closeKlondikeWinButton.addEventListener('click',()=>{hideGameModal(klondikeElements.klondikeWinMessage);});
    /* --- /캐주얼 게임즈 모달의 JavaScript 부분 --- */

    /* --- 앱 창(모달-윈도우) 열기/닫기 통합 연결 ---
       이전에는 앱마다 openXxxBtn / closeXxxBtn 변수를 일일이 선언하고 각각
       addEventListener를 붙였는데, 그 과정에서 캐주얼 게임즈의 "닫기" 버튼이
       실수로 "열기" 버튼과 같은 id(openCasualGamesModalBtn)를 다시 조회하는
       버그가 있었음. 그 결과 모달 안의 × 닫기 버튼은 아무 동작도 하지 않고,
       사이드바의 🎮 아이콘을 클릭하면 열기/닫기 핸들러가 동시에 같은 엘리먼트에
       걸려 모달이 열렸다가 바로 닫히는 것처럼 보였다.
       아래처럼 설정 배열 + 한 곳의 등록 함수로 묶으면, id를 잘못 짝지어도
       콘솔에 경고가 찍히고(아래 registerAppWindows 참고) 같은 실수가
       조용히 묻히지 않는다. */
    const APP_WINDOWS = [
        { name: '계산기', openBtnId: 'openCalculatorModalBtn', closeBtnId: 'closeCalculatorModalBtn', modal: calculatorModal, onOpen: Calculator.resetUI },
        { name: '달력', openBtnId: 'openCalendarModalBtn', closeBtnId: 'closeCalendarModalBtn', modal: calendarModal, onOpen: renderCalendar },
        { name: '워드프로세서', openBtnId: 'openWordProcessorModalBtn', closeBtnId: 'closeWordProcessorModalBtn', modal: wordProcessorModal },
        { name: '스프레드시트', openBtnId: 'openSpreadsheetModalBtn', closeBtnId: 'closeSpreadsheetModalBtn', modal: spreadsheetModal, onOpen: spreadsheetModalInit },
        { name: '캐주얼 게임즈', openBtnId: 'openCasualGamesModalBtn', closeBtnId: 'closeCasualGamesModalBtn', modal: casualGamesModal },
        { name: '도움말', openBtnId: 'openHelpModalBtn', closeBtnId: 'closeHelpModalBtn', modal: helpModal },
    ];

    const registerAppWindows = (configs) => {
        configs.forEach((cfg) => {
            const openBtn = document.getElementById(cfg.openBtnId);
            const closeBtn = document.getElementById(cfg.closeBtnId);
            if (!openBtn || !closeBtn || !cfg.modal) {
                console.warn(`[모달 연결 실패] ${cfg.name}: open=${!!openBtn}, close=${!!closeBtn}, modal=${!!cfg.modal}`);
                return;
            }
            cfg.openBtn = openBtn; // Esc 키 처리에서 dock 표시를 끄기 위해 보관
            openBtn.addEventListener('click', () => {
                openAppModal(cfg.modal);
                setDockActive(openBtn, true);
                if (cfg.onOpen) cfg.onOpen();
                // 메뉴바 앱 이름 갱신
                const titleEl = cfg.modal.querySelector && cfg.modal.querySelector('.titlebar-title');
                if (titleEl && window.OSShell) window.OSShell.setActiveApp(titleEl.textContent);
            });
            closeBtn.addEventListener('click', () => {
                closeAppModal(cfg.modal);
                setDockActive(openBtn, false);
            });
        });
    };
    registerAppWindows(APP_WINDOWS);
            window._APP_WINDOWS = APP_WINDOWS; // OS Shell에서 참조

    // Esc 키로 가장 앞에 떠 있는 앱 창을 닫을 수 있도록 (키보드 접근성 개선)
    document.addEventListener('keydown', (e) => {
        if (e.key !== 'Escape') return;
        const openConfigs = APP_WINDOWS.filter((cfg) => cfg.modal.classList.contains('active'));
        if (openConfigs.length === 0) return;
        const topConfig = openConfigs.reduce((top, cfg) =>
            (parseInt(cfg.modal.style.zIndex || '0', 10) > parseInt(top.modal.style.zIndex || '0', 10) ? cfg : top));
        closeAppModal(topConfig.modal);
        setDockActive(topConfig.openBtn, false);
    });

/* ────────────────────────────────────────────
   OS Shell: 메뉴바 · 신호등 버튼 · 컨텍스트 메뉴
   ──────────────────────────────────────────── */

const OSShell = (() => {

    // ── 메뉴바 시계 ──
    const clockEl  = document.getElementById('menubar-clock');
    const dotEl    = document.getElementById('menubar-online-dot');
    const activeAppEl = document.getElementById('menubar-active-app');

    function updateClock() {
        const now = new Date();
        const d = now.toLocaleDateString('ko-KR', { month: 'short', day: 'numeric', weekday: 'short' });
        const t = now.toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' });
        if (clockEl) clockEl.textContent = `${d}  ${t}`;
    }
    updateClock();
    setInterval(updateClock, 1000);

    // ── 온라인 상태 표시 ──
    function updateOnline() {
        if (!dotEl) return;
        if (navigator.onLine) {
            dotEl.classList.remove('offline');
            dotEl.title = '온라인';
        } else {
            dotEl.classList.add('offline');
            dotEl.title = '오프라인';
        }
    }
    updateOnline();
    window.addEventListener('online',  updateOnline);
    window.addEventListener('offline', updateOnline);

    // ── 포커스된 앱 이름 표시 ──
    function setActiveApp(name) {
        if (activeAppEl) activeAppEl.textContent = name || '';
    }

    // ── 신호등 버튼 핸들러 ──
    function handleTitlebarClick(e) {
        const btn = e.target.closest('.titlebar-btn');
        if (!btn) return;
        const action = btn.dataset.action;
        const targetId = btn.dataset.target;
        const win = document.getElementById(targetId);
        if (!win) return;

        if (action === 'close') {
            // APP_WINDOWS 설정에서 해당 모달의 openBtn 찾아 dock 해제
            const cfg = window._APP_WINDOWS && window._APP_WINDOWS.find(c => c.modal === win);
            if (cfg) {
                closeAppModal(win);
                setDockActive(cfg.openBtn, false);
                setActiveApp('');
            }
        } else if (action === 'minimize') {
            win.classList.add('minimized');
            // dock dot는 유지(실행중), app-active 클래스도 유지
            setActiveApp('');
        } else if (action === 'maximize') {
            if (win.classList.contains('maximized')) {
                win.classList.remove('maximized');
            } else {
                win.classList.remove('minimized');
                win.classList.add('maximized');
            }
        }
    }
    document.addEventListener('click', handleTitlebarClick);

    // ── 창 포커스 시 메뉴바 앱 이름 갱신 ──
    document.addEventListener('mousedown', e => {
        const win = e.target.closest('.modal-window');
        if (!win) return;
        const titleEl = win.querySelector('.titlebar-title');
        if (titleEl) setActiveApp(titleEl.textContent);
        // 최소화 상태 창을 클릭하면 복원 (Dock 아이콘 재클릭 시 대비)
        if (win.classList.contains('minimized')) {
            win.classList.remove('minimized');
        }
    });

    // ── 컨텍스트 메뉴 ──
    const ctxMenu = document.getElementById('context-menu');

    function hideCtx() {
        if (ctxMenu) ctxMenu.classList.remove('visible');
    }

    if (ctxMenu) {
        // 바탕화면(main-content) 우클릭 시 열기
        document.getElementById('main-content').addEventListener('contextmenu', e => {
            e.preventDefault();
            const x = Math.min(e.clientX, window.innerWidth  - ctxMenu.offsetWidth  - 8);
            const y = Math.min(e.clientY, window.innerHeight - ctxMenu.offsetHeight - 8);
            ctxMenu.style.left = x + 'px';
            ctxMenu.style.top  = y + 'px';
            ctxMenu.classList.add('visible');
        });

        // 컨텍스트 메뉴 항목 클릭
        ctxMenu.addEventListener('click', e => {
            const item = e.target.closest('.ctx-item');
            if (!item) return;
            hideCtx();
            const actionMap = {
                'open-calculator': 'openCalculatorModalBtn',
                'open-calendar':   'openCalendarModalBtn',
                'open-word':       'openWordProcessorModalBtn',
                'open-sheet':      'openSpreadsheetModalBtn',
                'open-games':      'openCasualGamesModalBtn',
                'open-help':       'openHelpModalBtn',
            };
            const btnId = actionMap[item.dataset.action];
            if (btnId) document.getElementById(btnId)?.click();
        });

        // 바깥 클릭 시 닫기
        document.addEventListener('click',       hideCtx);
        document.addEventListener('contextmenu', e => { if (!e.target.closest('#main-content')) hideCtx(); });
        document.addEventListener('keydown',     e => { if (e.key === 'Escape') hideCtx(); });
    }

    // ── Dock 아이콘 클릭: 최소화 상태면 복원 ──
    document.getElementById('sidebar').addEventListener('click', e => {
        const btn = e.target.closest('.modal-open-button');
        if (!btn) return;
        // APP_WINDOWS에서 해당 버튼의 모달 찾기
        const cfg = window._APP_WINDOWS && window._APP_WINDOWS.find(c => c.openBtn === btn);
        if (!cfg) return;
        if (cfg.modal.classList.contains('minimized')) {
            cfg.modal.classList.remove('minimized');
            bringModalToFront(cfg.modal);
            const titleEl = cfg.modal.querySelector('.titlebar-title');
            if (titleEl) setActiveApp(titleEl.textContent);
            e.stopPropagation(); // openBtn 의 원래 click도 발생하므로 중복 방지 불필요
        }
    }, true); // capture phase로 먼저 잡기

    window.OSShell = { setActiveApp };
    return { setActiveApp };
})();

});