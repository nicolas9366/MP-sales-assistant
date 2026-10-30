'use strict';
(() => {
  const data = window.BONYU_CONTENT || [];
  const grid = document.querySelector('#content-grid');
  if (!grid) return;
  const search = document.querySelector('#content-search');
  const dialog = document.querySelector('#article-dialog');
  let category = '全部';
  let opener = null;
  const node = (tag, text, className) => { const el = document.createElement(tag); if (text != null) el.textContent = text; if (className) el.className = className; return el; };
  const openArticle = (item, trigger) => {
    opener = trigger;
    document.querySelector('#article-title').textContent = item.title;
    document.querySelector('#article-meta').textContent = `${item.category} / ${item.type} / ${item.date}`;
    document.querySelector('#article-intro').textContent = item.summary;
    const body = document.querySelector('#article-body'); body.replaceChildren();
    for (const section of item.sections) {
      body.append(node('h3', section.heading));
      for (const p of section.paragraphs || []) body.append(node('p', p));
      if (section.items) { const list = node('ul'); section.items.forEach(text => list.append(node('li', text))); body.append(list); }
    }
    if (!dialog.open) dialog.showModal();
    dialog.scrollTop = 0;
  };
  const render = () => {
    const query = search.value.trim().toLocaleLowerCase();
    const filtered = data.filter(item => (category === '全部' || item.category === category) && `${item.title} ${item.summary} ${item.category} ${item.sections.map(s => [s.heading,...(s.paragraphs || []),...(s.items || [])].join(' ')).join(' ')}`.toLocaleLowerCase().includes(query));
    grid.replaceChildren();
    for (const item of filtered) {
      const card = node('article', null, 'content-card');
      const art = node('div', null, 'card-art'); art.setAttribute('aria-hidden', 'true'); art.style.setProperty('--art-bg', item.colors[0]); art.style.setProperty('--art-ink', item.colors[1]);
      const strong = node('strong'); strong.append(document.createTextNode(item.art[0]), node('br'), document.createTextNode(item.art[1])); art.append(strong);
      const meta = node('div', null, 'card-meta'); const time = node('time', item.date); time.dateTime = item.date; meta.append(node('span', `${item.category} · ${item.type}`, 'category'), time);
      const title = node('h3'); const link = node('a', item.title); link.href = `#article=${item.id}`; title.append(link);
      const read = node('a', '阅读全文', 'read-link'); read.href = link.href;
      for (const trigger of [link, read]) trigger.addEventListener('click', event => { event.preventDefault(); history.pushState(null, '', `#article=${item.id}`); openArticle(item, trigger); });
      card.append(art, meta, title, node('p', item.summary), read); grid.append(card);
    }
    document.querySelector('#result-count').textContent = `${filtered.length} 篇内容`;
    document.querySelector('#empty-state').hidden = filtered.length > 0;
  };
  const syncHash = () => {
    const id = location.hash.startsWith('#article=') ? location.hash.slice(9) : '';
    const item = data.find(item => item.id === id);
    if (item) openArticle(item, null); else if (dialog.open) dialog.close();
  };
  document.querySelectorAll('[data-filter]').forEach(button => button.addEventListener('click', () => { category = button.dataset.filter; document.querySelectorAll('[data-filter]').forEach(other => other.setAttribute('aria-pressed', String(other === button))); render(); }));
  search.addEventListener('input', render);
  document.querySelector('#close-article').addEventListener('click', () => dialog.close());
  dialog.addEventListener('close', () => { if (location.hash.startsWith('#article=')) history.replaceState(null, '', '#content'); if (opener && opener.isConnected) opener.focus(); });
  window.addEventListener('popstate', syncHash); window.addEventListener('hashchange', syncHash);
  render(); syncHash();
})();
