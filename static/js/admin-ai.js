/* ASKme Analytics — admin-only AI assistant widget */

(function () {
  const root = document.createElement('div');
  root.id = 'adminai';
  root.innerHTML = `
    <button class="adminai-toggle" aria-label="Open analytics AI">
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M12 2a4 4 0 0 0-4 4v2a4 4 0 0 0 8 0V6a4 4 0 0 0-4-4z"/>
        <path d="M4 14a4 4 0 0 1 4-4h8a4 4 0 0 1 4 4v1a7 7 0 0 1-7 7h-2a7 7 0 0 1-7-7v-1z"/>
        <circle cx="9" cy="6" r="0.5" fill="currentColor"/>
        <circle cx="15" cy="6" r="0.5" fill="currentColor"/>
      </svg>
    </button>

    <div class="adminai-panel" hidden>
      <header class="adminai-head">
        <div>
          <strong>ASKme Analytics</strong>
          <span>Admin insights &middot; live data</span>
        </div>
        <button class="adminai-close" aria-label="Close">&times;</button>
      </header>

      <div class="adminai-body" id="adminai-body">
        <div class="adminai-msg bot">
          <p>Hi Admin. Ask me anything about members, inventory, prints, attendance, or events.</p>
          <p class="adminai-hint">Try:</p>
          <div class="adminai-suggestions">
            <button class="adminai-suggest" data-q="Give me a quick dashboard summary">Dashboard summary</button>
            <button class="adminai-suggest" data-q="Any print requests pending approval?">Pending prints?</button>
            <button class="adminai-suggest" data-q="What items are overdue?">Overdue items?</button>
            <button class="adminai-suggest" data-q="How many active members do we have, broken down by role?">Active member count</button>
            <button class="adminai-suggest" data-q="Who hasn't logged in in 30 days?">Inactive members</button>
            <button class="adminai-suggest" data-q="What are the 5 most-checked-out items?">Top checked-out items</button>
          </div>
        </div>
      </div>

      <form class="adminai-form" id="adminai-form">
        <input type="text" id="adminai-input" placeholder="Ask about your data..." autocomplete="off" />
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

  const toggle = root.querySelector('.adminai-toggle');
  const panel  = root.querySelector('.adminai-panel');
  const close  = root.querySelector('.adminai-close');
  const body   = root.querySelector('#adminai-body');
  const form   = root.querySelector('#adminai-form');
  const input  = root.querySelector('#adminai-input');

  let history = [];
  let busy = false;

  toggle.addEventListener('click', () => {
    panel.hidden = false;
    toggle.classList.add('open');
    setTimeout(() => input.focus(), 100);
  });
  close.addEventListener('click', () => {
    panel.hidden = true;
    toggle.classList.remove('open');
  });

  function addUser(text) {
    const el = document.createElement('div');
    el.className = 'adminai-msg user';
    el.textContent = text;
    body.appendChild(el);
    scroll();
  }

  function addBot(text, toolCalls) {
    const el = document.createElement('div');
    el.className = 'adminai-msg bot';

    // Render markdown-ish (bold, code, line breaks, bullets)
    const html = text
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
      .replace(/`([^`]+)`/g, '<code>$1</code>')
      .replace(/^- (.+)$/gm, '• $1')
      .replace(/\n/g, '<br>');
    el.innerHTML = html;

    if (toolCalls && toolCalls.length) {
      const badge = document.createElement('div');
      badge.className = 'adminai-toolbadge';
      badge.textContent = `${toolCalls.length} data query${toolCalls.length > 1 ? 'ies' : ''}: ` +
        toolCalls.map(tc => tc.name).join(', ');
      el.appendChild(badge);
    }
    body.appendChild(el);
    scroll();
  }

  function addTyping() {
    const el = document.createElement('div');
    el.className = 'adminai-msg bot adminai-typing';
    el.innerHTML = '<span></span><span></span><span></span>';
    el.id = 'adminai-typing';
    body.appendChild(el);
    scroll();
    return el;
  }

  function scroll() {
    body.scrollTop = body.scrollHeight;
  }

  async function send(q) {
    if (busy || !q.trim()) return;
    busy = true;
    input.disabled = true;
    addUser(q);
    const typingEl = addTyping();

    try {
      const resp = await fetch('/api/ask', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message: q, history }),
      });
      typingEl.remove();

      if (!resp.ok) {
        const err = await resp.json().catch(() => ({ error: 'request failed' }));
        addBot(`Error: ${err.error || resp.status}`);
        return;
      }
      const data = await resp.json();
      addBot(data.reply || '(no answer)', data.tool_calls);

      // Update local conversation history (store just user+assistant text for context)
      history.push({ role: 'user', content: q });
      history.push({ role: 'assistant', content: data.reply });
      // Keep history reasonable
      if (history.length > 20) history = history.slice(-20);
    } catch (e) {
      typingEl.remove();
      addBot('Network error. Check the server is running.');
    } finally {
      busy = false;
      input.disabled = false;
      input.focus();
    }
  }

  form.addEventListener('submit', e => {
    e.preventDefault();
    const q = input.value.trim();
    if (!q) return;
    input.value = '';
    send(q);
  });

  // Suggestion chips
  root.addEventListener('click', e => {
    const btn = e.target.closest('.adminai-suggest');
    if (!btn) return;
    send(btn.dataset.q);
  });
})();
