/* 사이트 공통 상단 메뉴 (통계 분석 ↔ 모델 성능지표)
   각 페이지 맨 아래에 <script src="…/nav.js"></script> 한 줄만 넣으면 된다.
   메뉴 항목을 바꾸거나 새 도구를 추가할 때는 아래 ITEMS 만 고치면 모든 페이지에 반영된다. */
(function () {
  'use strict';
  if (document.getElementById('site-nav')) return;
  var script = document.currentScript;
  if (!script || !script.src) return;

  // nav.js 가 놓인 폴더 = 사이트 루트(docs/). 어느 깊이의 페이지에서 불러와도 같은 주소가 나온다.
  var base = new URL('.', script.src).href;
  // file:// 로 직접 열면 폴더 주소는 열리지 않으므로 index.html 까지 적는다.
  var tail = location.protocol === 'file:' ? 'index.html' : '';

  var ITEMS = [
    { label: '통계 분석', href: base + tail },
    { label: '모델 성능지표', href: base + 'metrics/' + tail }
  ];

  function plain(u) { return u.split('#')[0].split('?')[0].replace(/index\.html$/, ''); }
  var here = plain(location.href);

  var css =
    '.sn-bar{box-sizing:border-box;width:100%;margin:0;padding:0;background:#fff;border-bottom:1px solid #d9dde5;' +
      'font:600 15px/1 "Pretendard","Noto Sans KR","Apple SD Gothic Neo","Malgun Gothic",system-ui,sans-serif}' +
    '.sn-inner{box-sizing:border-box;max-width:1320px;margin:0 auto;padding:0 12px;display:flex;gap:4px}' +
    '.sn-link{box-sizing:border-box;display:inline-flex;align-items:center;min-height:46px;padding:0 18px;' +
      'color:#4f5566;text-decoration:none;border-bottom:3px solid transparent;margin-bottom:-1px}' +
    '.sn-link:hover{color:#1B2A4A;background:#f1f3f8}' +
    '.sn-link[aria-current="page"]{color:#1B2A4A;border-bottom-color:#1B2A4A}' +
    '.sn-link:focus-visible{outline:3px solid #1c4fd8;outline-offset:-3px}' +
    '@media (max-width:480px){.sn-inner{padding:0 4px}.sn-link{padding:0 14px}}' +
    '@media print{.sn-bar{display:none}}';
  var style = document.createElement('style');
  style.id = 'site-nav-style';
  style.textContent = css;
  document.head.appendChild(style);

  var nav = document.createElement('nav');
  nav.id = 'site-nav';
  nav.className = 'sn-bar';
  nav.setAttribute('aria-label', '사이트 메뉴');
  var inner = document.createElement('div');
  inner.className = 'sn-inner';
  ITEMS.forEach(function (it) {
    var a = document.createElement('a');
    a.className = 'sn-link';
    a.href = it.href;
    a.textContent = it.label;
    if (plain(it.href) === here) a.setAttribute('aria-current', 'page');
    inner.appendChild(a);
  });
  nav.appendChild(inner);
  document.body.insertBefore(nav, document.body.firstChild);
})();
