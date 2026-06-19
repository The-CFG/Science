// ═══════════════════════════════════════════════════════════
//  htmlblast.js  –  TheCFG Office / HTMLBlast module
//  (HTML 코드 에디터 + 실시간 미리보기)
// ═══════════════════════════════════════════════════════════

const HtmlBlast = (() => {
  // ── 태그 삽입 템플릿 정의 ─────────────────────────────────
  // cursor: 삽입 후 커서가 위치할 오프셋
  //   - 'end'        : textToInsert 끝으로 이동
  //   - 'innerStart'  : 첫 번째 '>' 바로 뒤 (태그 내부 시작)
  const TAG_TEMPLATES = {
    p:      { text: `<p>새 단락 텍스트</p>`,                                                          cursor: 'innerStart' },
    h1:     { text: `<h1>헤더 1 제목</h1>`,                                                           cursor: 'innerStart' },
    h2:     { text: `<h2>헤더 2 제목</h2>`,                                                           cursor: 'innerStart' },
    h3:     { text: `<h3>헤더 3 제목</h3>`,                                                           cursor: 'innerStart' },
    h4:     { text: `<h4>헤더 4 제목</h4>`,                                                           cursor: 'innerStart' },
    h5:     { text: `<h5>헤더 5 제목</h5>`,                                                           cursor: 'innerStart' },
    h6:     { text: `<h6>헤더 6 제목</h6>`,                                                           cursor: 'innerStart' },
    strong: { text: `<strong>강조된 텍스트</strong>`,                                                  cursor: 'innerStart' },
    div:    { text: `<div>새로운 div 내용</div>`,                                                      cursor: 'innerStart' },
    span:   { text: `<span>새로운 span 내용</span>`,                                                   cursor: 'innerStart' },
    a:      { text: `<a href="https://example.com">링크 텍스트</a>`,                                   cursor: 'innerStart' },
    img:    { text: `<img src="https://placehold.co/150x100?text=이미지" alt="이미지 설명">`,          cursor: 'end' },
    ul:     { text: `<ul>\n    <li>항목 1</li>\n    <li>항목 2</li>\n</ul>`,                          cursor: 'innerStart' },
    ol:     { text: `<ol>\n    <li>항목 1</li>\n    <li>항목 2</li>\n</ol>`,                          cursor: 'innerStart' },
    li:     { text: `<li>목록 항목</li>`,                                                              cursor: 'innerStart' },
    nav:    { text: `<nav>\n    <!-- 내비게이션 링크 -->\n    <ul>\n        <li><a href="#">홈</a></li>\n        <li><a href="#">소개</a></li>\n    </ul>\n</nav>`, cursor: 'innerStart' },
    header: { text: `<header>\n    <h1>사이트 제목</h1>\n    <!-- 로고, 메뉴 등 -->\n</header>`,      cursor: 'innerStart' },
    footer: { text: `<footer>\n    <p>&copy; 2024 내 웹사이트. 모든 권리 보유.</p>\n</footer>`,      cursor: 'innerStart' },
    main:   { text: `<main>\n    <!-- 페이지의 주요 콘텐츠 -->\n    <p>이곳에 내용을 채워보세요.</p>\n</main>`, cursor: 'innerStart' },
    input:  { text: `<input type="text" placeholder="텍스트 입력">`,                                  cursor: 'end' },
    button: { text: `<button>클릭하세요</button>`,                                                     cursor: 'innerStart' },
    form:   { text: `<form>\n    <label for="name">이름:</label>\n    <input type="text" id="name" name="name">\n    <br>\n    <button type="submit">제출</button>\n</form>`, cursor: 'innerStart' },
    table:  { text: `<table border="1" style="width:100%;">\n    <thead>\n        <tr><th>헤더 1</th><th>헤더 2</th></tr>\n    </thead>\n    <tbody>\n        <tr><td>데이터 1</td><td>데이터 2</td></tr>\n        <tr><td>데이터 3</td><td>데이터 4</td></tr>\n    </tbody>\n</table>`, cursor: 'innerStart' },
  };

  const INITIAL_HTML = `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>나만의 멋진 웹페이지</title>
    <style>
        body {
            font-family: sans-serif;
            padding: 20px;
            background-color: #f4f4f4;
            color: #333;
        }
        h1 { color: #2c3e50; }
        p  { line-height: 1.6; }
        strong { color: #e74c3c; }
    </style>
</head>
<body>
    <h1>환영합니다!</h1>
    <p>이것은 <b>당신의 HTML 에디터</b>입니다. 위 버튼을 눌러 요소를 삽입하거나, 직접 코드를 작성해보세요.</p>
    <p>강조하고 싶은 텍스트는 <strong>strong</strong>을 사용해 보세요.</p>
    <p>미리보기 상단에 이 페이지의 <b>제목</b>이 표시되는 것을 확인해보세요!</p>
    <div>이것은 div 태그입니다. 여러 요소를 묶을 수 있어요.</div>
    <ul>
        <li>새로운</li>
        <li>목록</li>
        <li>항목</li>
    </ul>
</body>
</html>`;

  // ── DOM ────────────────────────────────────────────────────
  let editor, previewFrame, previewTitle, tagButtonsContainer, charCountEl;

  // ── Init ───────────────────────────────────────────────────
  function init() {
    editor               = document.getElementById('htmlblast-code-editor');
    previewFrame         = document.getElementById('htmlblast-preview-frame');
    previewTitle         = document.getElementById('htmlblast-preview-title');
    tagButtonsContainer  = document.getElementById('htmlblast-tag-buttons');
    charCountEl          = document.getElementById('htmlblast-char-count');

    if (!editor || !previewFrame || !previewTitle || !tagButtonsContainer) {
      console.error('[HTMLBlast] 필수 DOM 요소를 찾을 수 없습니다.');
      return;
    }

    // 초기 코드 설정 및 미리보기 렌더링
    editor.value = INITIAL_HTML;
    renderPreview();
    updateStatus();

    editor.addEventListener('input', updateStatus);

    // 태그 버튼 이벤트 등록
    tagButtonsContainer.querySelectorAll('button[data-tag]').forEach(button => {
      button.addEventListener('click', () => {
        insertTag(button.dataset.tag);
        renderPreview();
        updateStatus();
      });
    });
  }

  // ── Status bar ─────────────────────────────────────────────
  function updateStatus() {
    if (charCountEl) charCountEl.textContent = editor.value.length.toLocaleString();
  }

  // ── Preview ────────────────────────────────────────────────
  function renderPreview() {
    const code = editor.value;
    previewFrame.srcdoc = code;

    const titleMatch = code.match(/<\s*title\s*>(.*?)<\s*\/\s*title\s*>/is);
    previewTitle.textContent = (titleMatch && titleMatch[1])
      ? titleMatch[1].trim()
      : '미리보기';
  }

  function run() {
    renderPreview();
  }

  // ── 태그 삽입 ──────────────────────────────────────────────
  function insertTag(tagName) {
    const template = TAG_TEMPLATES[tagName];
    if (!template) return;

    const start    = editor.selectionStart;
    const end      = editor.selectionEnd;
    const selected = editor.value.substring(start, end);
    const before   = editor.value.substring(0, start);
    const after    = editor.value.substring(end);

    let textToInsert = template.text;

    // 텍스트가 선택된 경우, 선택 텍스트를 태그로 감싸기
    if (selected.length > 0 && !['img', 'input'].includes(tagName)) {
      const openTag  = textToInsert.match(/^(<[^>]+>)/)?.[1] ?? '';
      const closeTag = textToInsert.match(/(<\/[^>]+>)$/)?.[1] ?? '';
      if (openTag && closeTag) {
        textToInsert = openTag + selected + closeTag;
      }
    }

    editor.value = before + textToInsert + after;

    // 커서 위치 결정
    let newCursor;
    if (template.cursor === 'end') {
      newCursor = start + textToInsert.length;
    } else {
      // 'innerStart': 첫 번째 닫는 꺾쇠(>) 바로 뒤
      const firstClose = textToInsert.indexOf('>');
      newCursor = start + (firstClose >= 0 ? firstClose + 1 : textToInsert.length);
    }

    editor.selectionStart = editor.selectionEnd = newCursor;
    editor.focus();
  }

  // ── 저장 ───────────────────────────────────────────────────
  function save() {
    const code = editor.value;
    if (!code.trim()) {
      Word.showMsg('저장할 내용이 없습니다.');
      return;
    }

    const blob = new Blob([code], { type: 'text/html' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = 'my_webpage.html';

    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    Word.showMsg('"my_webpage.html"(으)로 저장되었습니다.');
  }

  // ── 불러오기 ───────────────────────────────────────────────
  function load() {
    const input  = document.createElement('input');
    input.type   = 'file';
    input.accept = '.html,.htm,.txt';

    // 선택 여부와 무관하게 항상 DOM 정리
    const cleanup = () => {
      if (input.parentNode) document.body.removeChild(input);
    };

    input.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file) { cleanup(); return; }

      const reader   = new FileReader();
      reader.onload  = (event) => {
        editor.value = event.target.result;
        renderPreview();
        updateStatus();
        cleanup();
        Word.showMsg(`"${file.name}" 파일이 로드되었습니다.`);
      };
      reader.onerror = () => {
        Word.showMsg('파일을 읽는 중 오류가 발생했습니다.');
        cleanup();
      };
      reader.readAsText(file);
    });

    // 파일 선택 창을 닫았을 때(cancel) 정리
    // focus 이벤트는 대화상자가 닫힌 직후 window로 돌아올 때 발생
    window.addEventListener('focus', cleanup, { once: true });

    input.style.display = 'none';
    document.body.appendChild(input);
    input.click();
  }

  // ── Public API ─────────────────────────────────────────────
  return { init, run, save, load };
})();