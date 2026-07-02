let personas = [];
let activeIdx = null;
let rarityFilterVal = null;
let sinnerFilterVal = null;

const fileInput = document.getElementById('fileInput');
const fileDrop = document.getElementById('fileDrop');
const personaListEl = document.getElementById('personaList');
const searchInput = document.getElementById('searchInput');
const rarityFilterEl = document.getElementById('rarityFilter');
const sinnerFilterEl = document.getElementById('sinnerFilter');
const mainPanel = document.getElementById('mainPanel');

fileInput.addEventListener('change', e => loadFiles(e.target.files));
fileDrop.addEventListener('dragover', e => { e.preventDefault(); fileDrop.classList.add('dragover'); });
fileDrop.addEventListener('dragleave', () => fileDrop.classList.remove('dragover'));
fileDrop.addEventListener('drop', e => {
  e.preventDefault();
  fileDrop.classList.remove('dragover');
  loadFiles(e.dataTransfer.files);
});

function loadFiles(fileList){
  const files = Array.from(fileList).filter(f => f.name.endsWith('.json'));
  let pending = files.length;
  if(pending === 0) return;
  files.forEach(file => {
    const reader = new FileReader();
    reader.onload = ev => {
      try{
        const data = JSON.parse(ev.target.result);
        personas.push(data);
      }catch(err){
        console.error('파싱 실패:', file.name, err);
      }
      pending--;
      if(pending === 0) onAllLoaded();
    };
    reader.readAsText(file, 'UTF-8');
  });
}

function onAllLoaded(){
  personas.sort((a,b) => (a.sinner||'').localeCompare(b.sinner||'') || (a.rarity||0)-(b.rarity||0));
  buildFilters();
  renderList();
}

function buildFilters(){
  const rarities = [...new Set(personas.map(p => p.rarity).filter(Boolean))].sort();
  rarityFilterEl.innerHTML = rarities.map(r =>
    `<button class="chip" data-rarity="${r}">${r}성</button>`
  ).join('');
  rarityFilterEl.querySelectorAll('.chip').forEach(btn => {
    btn.addEventListener('click', () => {
      const val = btn.dataset.rarity;
      rarityFilterVal = (rarityFilterVal === val) ? null : val;
      buildFilters(); sinnerChipsRefresh(); renderList();
    });
    if(String(rarityFilterVal) === btn.dataset.rarity) btn.classList.add('active');
  });

  sinnerChipsRefresh();
}

function sinnerChipsRefresh(){
  const sinners = [...new Set(personas.map(p => p.sinner).filter(Boolean))];
  sinnerFilterEl.innerHTML = sinners.map(s =>
    `<button class="chip" data-sinner="${s}">${s}</button>`
  ).join('');
  sinnerFilterEl.querySelectorAll('.chip').forEach(btn => {
    btn.addEventListener('click', () => {
      const val = btn.dataset.sinner;
      sinnerFilterVal = (sinnerFilterVal === val) ? null : val;
      sinnerChipsRefresh(); renderList();
    });
    if(sinnerFilterVal === btn.dataset.sinner) btn.classList.add('active');
  });
}

searchInput.addEventListener('input', renderList);

function renderList(){
  const q = searchInput.value.trim().toLowerCase();
  const filtered = personas
    .map((p, idx) => ({p, idx}))
    .filter(({p}) => {
      if(rarityFilterVal && String(p.rarity) !== rarityFilterVal) return false;
      if(sinnerFilterVal && p.sinner !== sinnerFilterVal) return false;
      if(q && !(`${p.name} ${p.sinner}`.toLowerCase().includes(q))) return false;
      return true;
    });

  if(filtered.length === 0){
    personaListEl.innerHTML = `<div class="empty-sidebar">조건에 맞는 인격이 없습니다.</div>`;
    return;
  }

  personaListEl.innerHTML = filtered.map(({p, idx}) => `
    <div class="persona-item ${idx===activeIdx?'active':''}" data-idx="${idx}">
      <div class="pname">${escapeHtml(p.name || '이름 없음')}</div>
      <div class="pmeta">
        ${rarityIconHtml(p.rarity)}
        ${escapeHtml(p.sinner||'')}
      </div>
    </div>
  `).join('');

  personaListEl.querySelectorAll('.persona-item').forEach(el => {
    el.addEventListener('click', () => {
      activeIdx = Number(el.dataset.idx);
      renderList();
      renderDetail(personas[activeIdx]);
    });
  });
}

function escapeHtml(str){
  return String(str).replace(/[&<>"']/g, c => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[c]));
}

// 등급(rarity) 뱃지 이미지를 표시합니다. assets/IDNumber{n}.webp 파일이 있으면 사용하고,
// 없는 등급은 자동으로 기본 점(dot) 표시로 대체됩니다.
function rarityIconHtml(rarity){
  if(!rarity) return '<span class="rarity-dot"></span>';
  return `<img class="rarity-icon" src="assets/IDNumber${rarity}.webp" alt="${rarity}성" onerror="this.outerHTML='&lt;span class=&quot;rarity-dot&quot;&gt;&lt;/span&gt;'">`;
}

function sinColor(sin){
  return `var(--sin-${sin})` || 'var(--gold)';
}

function renderDetail(p){
  const resistOrder = ['참격','관통','타격'];
  const resistHtml = resistOrder.map(type => {
    const r = (p.resistance||{})[type];
    if(!r) return '';
    return `
      <div class="resist-row">
        <span>${type}</span>
        <span class="resist-tag ${r.level}">${r.level} ${r.multiplier||''}</span>
      </div>`;
  }).join('');

  const statsHtml = p.stats ? `
    <div class="stat-row"><span class="k">HP</span><span class="v">${escapeHtml(p.stats.maxLevel||'-')}</span></div>
    <div class="stat-row"><span class="k">속도</span><span class="v">${escapeHtml(p.stats.speed||'-')}</span></div>
    <div class="stat-row"><span class="k">방어력</span><span class="v">${escapeHtml(p.stats.hp||'-')}</span></div>
    <div class="stat-row"><span class="k">출시</span><span class="v">${escapeHtml(p.releaseDate||'-')}</span></div>
    <div class="stat-row"><span class="k">시즌</span><span class="v">${escapeHtml(p.season||'-')}</span></div>
  ` : '<div class="stat-row"><span class="k">데이터 없음</span></div>';

  const skillsHtml = (p.skills||[]).map(s => skillCard(s, false)).join('');
  const defenseHtml = p.defenseSkill ? skillCard(p.defenseSkill, true) : '';

  const passivesHtml = (p.passives||[]).map(pv => `
    <div class="passive-card">
      <div class="ptitle">
        <span class="pname">${escapeHtml(pv.name||'')}</span>
        <span class="ptype">${escapeHtml(pv.type||'')} · ${escapeHtml(pv.resonance||'')}</span>
      </div>
      <div class="pdesc">${escapeHtml(pv.description||'')}</div>
    </div>
  `).join('');

  const syncHtml = (p.syncUpgrades||[]).map(s => `
    <div class="sync-stage">
      <div class="snum">${s.stage}단계<br><span style="color:var(--ink-faint)">${escapeHtml(s.cost||'')}</span></div>
      <div class="seff">${(s.effects||[]).map(escapeHtml).join(' · ')}</div>
    </div>
  `).join('');

  const sanityHtml = p.sanity ? `
    <div style="font-family:var(--font-mono); font-size:12px; color:var(--ink-dim); margin-bottom:10px;">
      패닉 유형: <b style="color:var(--ink)">${escapeHtml(p.sanity.panicType||'-')}</b> — ${escapeHtml(p.sanity.panicEffect||'')}
    </div>
    <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px;">
      <div>
        <div style="font-size:11px; color:var(--sin-질투); font-family:var(--font-mono); margin-bottom:6px;">증가 조건</div>
        <ul class="sanity-list">${(p.sanity.increaseConditions||[]).map(c=>`<li>${escapeHtml(c)}</li>`).join('')}</ul>
      </div>
      <div>
        <div style="font-size:11px; color:var(--sin-분노); font-family:var(--font-mono); margin-bottom:6px;">감소 조건</div>
        <ul class="sanity-list">${(p.sanity.decreaseConditions||[]).map(c=>`<li>${escapeHtml(c)}</li>`).join('')}</ul>
      </div>
    </div>
  ` : '';

  mainPanel.innerHTML = `
    <div class="dossier-head">
      <div class="titles">
        <div class="sinner">${escapeHtml(p.sinner||'')}</div>
        <h2>${escapeHtml(p.name||'')}</h2>
        <div class="sub">
          <span>${escapeHtml(p.acquireMethod||'-')}</span>
          <span>${(p.keywords||[]).join(', ')}</span>
        </div>
      </div>
      <div>${rarityIconHtml(p.rarity)}</div>
    </div>

    <div class="grid-2">
      <div class="panel">
        <h3>스테이터스</h3>
        ${statsHtml}
      </div>
      <div class="panel">
        <h3>내성 정보</h3>
        ${resistHtml || '<div class="stat-row"><span class="k">데이터 없음</span></div>'}
      </div>
    </div>

    ${skillsHtml || defenseHtml ? `
    <section class="block">
      <h3>전투 스킬</h3>
      ${skillsHtml}
      ${defenseHtml}
    </section>` : ''}

    ${passivesHtml ? `
    <section class="block">
      <h3>패시브</h3>
      ${passivesHtml}
    </section>` : ''}

    ${syncHtml ? `
    <section class="block">
      <h3>동기화 강화</h3>
      ${syncHtml}
    </section>` : ''}

    ${sanityHtml ? `
    <section class="block">
      <h3>정신력 (패닉) 조건</h3>
      ${sanityHtml}
    </section>` : ''}
  `;
}

function skillCard(s, isDefense){
  const type = isDefense ? (s.defenseType||'-') : (s.attackType||'-');
  const sin = s.sinType || '';
  const coinEffects = (s.coinEffects||[]).map(ce => `
    <div class="coin-effect"><span class="cond">${escapeHtml(ce.condition||'')}</span>${escapeHtml(ce.effect||'')}</div>
  `).join('');

  return `
    <div class="skill-card ${isDefense?'defense':''}">
      <div class="skill-head">
        <span class="sname">${escapeHtml(s.name||'')}</span>
        ${sin ? `<span class="sin-tag" style="background:${sinColor(sin)}">${escapeHtml(sin)}</span>` : ''}
      </div>
      <div class="skill-stats">
        <span>${isDefense?'수비':'공격'} 유형 <b>${escapeHtml(type)}</b></span>
        <span>코인 수 <b>${escapeHtml(String(s.coinCount||'-'))}</b></span>
        <span>스킬 위력 <b>${escapeHtml(String(s.skillPower||'-'))}</b></span>
        <span>코인 위력 <b>${escapeHtml(String(s.coinPower||'-'))}</b></span>
        <span>가중치 <b>${escapeHtml(String(s.attackWeight||'-'))}</b></span>
      </div>
      ${coinEffects}
    </div>
  `;
}