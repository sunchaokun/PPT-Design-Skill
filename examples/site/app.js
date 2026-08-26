const DATA_URL = 'data/examples.json';
let projects = [];

async function loadProjects() {
  const response = await fetch(DATA_URL);
  projects = await response.json();
  if (document.getElementById('projects')) renderGallery();
  if (document.getElementById('viewer')) renderViewer();
}

function renderGallery() {
  document.getElementById('project-count').textContent = projects.length;
  document.getElementById('slide-count').textContent = projects.reduce((sum, p) => sum + p.slides.length, 0);
  const filters = ['全部', ...new Set(projects.map(p => p.category))];
  const filterNode = document.getElementById('filters');
  filterNode.innerHTML = filters.map((f, i) => `<button class="filter${i === 0 ? ' active' : ''}" data-filter="${f}">${f}</button>`).join('');
  filterNode.addEventListener('click', event => {
    const button = event.target.closest('.filter');
    if (!button) return;
    document.querySelectorAll('.filter').forEach(node => node.classList.toggle('active', node === button));
    renderCards(button.dataset.filter);
  });
  renderCards('全部');
}

function renderCards(category) {
  const list = category === '全部' ? projects : projects.filter(p => p.category === category);
  document.getElementById('projects').innerHTML = list.map((p, index) => `<a class="project-card" href="viewer.html?project=${encodeURIComponent(p.id)}"><div class="card-image"><img src="${p.path}/${p.slides[0].file}" alt="${p.title} 封面" loading="lazy"><div class="card-overlay"></div></div><div class="card-info"><div><h3>${p.title}</h3><p>${p.subtitle}</p></div><span class="card-number">${String(index + 1).padStart(2, '0')} / ${String(p.slides.length).padStart(2, '0')}</span></div></a>`).join('');
}

function renderViewer() {
  const id = new URLSearchParams(location.search).get('project');
  const project = projects.find(p => p.id === id) || projects[0];
  let current = 0;
  document.title = `${project.title} · PPT Design Skill`;
  document.getElementById('case-category').textContent = `${project.category} / ${project.year}`;
  document.getElementById('case-title').textContent = project.title;
  document.getElementById('case-description').textContent = project.description;
  document.getElementById('case-details').innerHTML = [['视觉方向', project.direction], ['页数', `${project.slides.length} 页`], ['生成模式', project.mode]].map(([k, v]) => `<div class="detail"><span>${k}</span><strong>${v}</strong></div>`).join('');
  document.getElementById('download-pptx').href = project.pptx;
  document.getElementById('download-pdf').href = project.pdf;
  const thumbs = document.getElementById('thumbs');
  thumbs.innerHTML = project.slides.map((slide, i) => `<button class="thumb${i === 0 ? ' active' : ''}" data-index="${i}" aria-label="第 ${i + 1} 页"><img src="${project.path}/${slide.file}" alt=""></button>`).join('');
  thumbs.addEventListener('click', event => { const button = event.target.closest('.thumb'); if (button) { current = Number(button.dataset.index); update(); }});
  document.getElementById('prev').onclick = () => { current = (current - 1 + project.slides.length) % project.slides.length; update(); };
  document.getElementById('next').onclick = () => { current = (current + 1) % project.slides.length; update(); };
  document.addEventListener('keydown', event => { if (event.key === 'ArrowLeft') document.getElementById('prev').click(); if (event.key === 'ArrowRight') document.getElementById('next').click(); });
  function update() { const slide = project.slides[current]; document.getElementById('current-slide').src = `${project.path}/${slide.file}`; document.getElementById('current-slide').alt = `${project.title} 第 ${current + 1} 页`; document.getElementById('slide-counter').textContent = `${String(current + 1).padStart(2, '0')} / ${String(project.slides.length).padStart(2, '0')}`; document.getElementById('slide-label').textContent = slide.label || ''; document.querySelectorAll('.thumb').forEach((node, i) => node.classList.toggle('active', i === current)); }
  update();
}

loadProjects().catch(error => { document.body.insertAdjacentHTML('beforeend', `<p style="padding:2rem;color:#b44c35">案例数据加载失败：${error.message}</p>`); });
