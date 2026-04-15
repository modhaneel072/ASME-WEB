/* ASKme — pre-programmed FAQ widget for the ASME public site */

(async function () {
  const FAQ_URL = '/static/js/faq-data.json';
  const data = await fetch(FAQ_URL).then(r => r.json());

  /* ── Build DOM ───────────────────────── */
  const root = document.createElement('div');
  root.id = 'faqbot';
  root.innerHTML = `
    <button class="faqbot-toggle" aria-label="Open chat">
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
      </svg>
    </button>

    <div class="faqbot-panel" hidden>
      <header class="faqbot-head">
        <div class="faqbot-head-text">
          <strong>ASKme</strong>
          <span>ASME at Iowa &middot; FAQ</span>
        </div>
        <button class="faqbot-close" aria-label="Close">&times;</button>
      </header>

      <div class="faqbot-body" id="faqbot-body">
        <div class="faqbot-msg bot">
          <p>Hi! I can answer common questions about ASME at Iowa. Pick a topic or type your question below.</p>
        </div>
        <div class="faqbot-chips" id="faqbot-chips"></div>
      </div>

      <form class="faqbot-form" id="faqbot-form">
        <input type="text" id="faqbot-input" placeholder="Ask a question..." autocomplete="off" />
        <button type="submit" aria-label="Send">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
            <line x1="22" y1="2" x2="11" y2="13"/>
            <polygon points="22 2 15 22 11 13 2 9 22 2"/>
          </svg>
        </button>
      </form>
    </div>
  `;
  document.body.appendChild(root);

  const toggle  = root.querySelector('.faqbot-toggle');
  const panel   = root.querySelector('.faqbot-panel');
  const close   = root.querySelector('.faqbot-close');
  const body    = root.querySelector('#faqbot-body');
  const chips   = root.querySelector('#faqbot-chips');
  const form    = root.querySelector('#faqbot-form');
  const input   = root.querySelector('#faqbot-input');

  /* ── Render category chips ──────────── */
  data.categories.forEach(cat => {
    const chip = document.createElement('button');
    chip.type = 'button';
    chip.className = 'faqbot-chip';
    chip.textContent = cat.label;
    chip.dataset.cat = cat.id;
    chip.addEventListener('click', () => showCategory(cat.id, cat.label));
    chips.appendChild(chip);
  });

  /* ── Open / close ──────────────────── */
  toggle.addEventListener('click', () => {
    panel.hidden = false;
    toggle.classList.add('open');
    setTimeout(() => input.focus(), 100);
  });
  close.addEventListener('click', () => {
    panel.hidden = true;
    toggle.classList.remove('open');
  });

  /* ── Helpers ───────────────────────── */
  function addUser(text) {
    const el = document.createElement('div');
    el.className = 'faqbot-msg user';
    el.innerHTML = `<p>${escapeHtml(text)}</p>`;
    body.insertBefore(el, chips);
    scroll();
  }

  function addBot(text, link, suggestions) {
    const el = document.createElement('div');
    el.className = 'faqbot-msg bot';
    let html = `<p>${text}</p>`;
    if (link) html += `<a class="faqbot-link" href="${link.url}">${link.text} &rarr;</a>`;
    el.innerHTML = html;
    body.insertBefore(el, chips);

    if (suggestions && suggestions.length) {
      const sgrp = document.createElement('div');
      sgrp.className = 'faqbot-suggestions';
      suggestions.forEach(q => {
        const s = document.createElement('button');
        s.type = 'button';
        s.className = 'faqbot-suggest';
        s.textContent = q.question;
        s.addEventListener('click', () => answer(q));
        sgrp.appendChild(s);
      });
      body.insertBefore(sgrp, chips);
    }
    scroll();
  }

  function escapeHtml(t) {
    return t.replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  }

  function scroll() {
    body.scrollTop = body.scrollHeight;
  }

  function showCategory(catId, catLabel) {
    addUser(catLabel);
    const items = data.faqs.filter(f => f.category === catId);
    addBot(`Pick a question about <strong>${catLabel.toLowerCase()}</strong>:`, null, items);
  }

  function answer(faq) {
    addUser(faq.question);
    setTimeout(() => addBot(faq.answer, faq.link), 300);
  }

  /* ── Search: fuzzy keyword match ───── */
  function search(q) {
    const query = q.toLowerCase().trim();
    if (!query) return [];

    const scored = data.faqs.map(f => {
      let score = 0;
      // direct question match
      if (f.question.toLowerCase().includes(query)) score += 10;
      // keyword matches
      f.keywords.forEach(k => {
        if (query.includes(k.toLowerCase())) score += 5;
        else if (k.toLowerCase().includes(query)) score += 2;
      });
      // word-by-word match in question
      query.split(/\s+/).forEach(word => {
        if (word.length < 3) return;
        if (f.question.toLowerCase().includes(word)) score += 2;
        f.keywords.forEach(k => { if (k.toLowerCase().includes(word)) score += 1; });
      });
      return { faq: f, score };
    });

    return scored.filter(s => s.score > 0).sort((a,b) => b.score - a.score).map(s => s.faq);
  }

  form.addEventListener('submit', e => {
    e.preventDefault();
    const q = input.value.trim();
    if (!q) return;
    addUser(q);
    input.value = '';

    const results = search(q);
    if (results.length === 0) {
      setTimeout(() => addBot("I don't have an answer for that yet. Try one of the topics above, or use the Contact page to reach a team lead.",
        { text: 'Contact us', url: '/contact' }), 300);
    } else if (results.length === 1) {
      setTimeout(() => addBot(results[0].answer, results[0].link), 300);
    } else {
      const top = results[0];
      const others = results.slice(1, 4);
      setTimeout(() => addBot(top.answer, top.link, others.length ? others : null), 300);
    }
  });
})();
