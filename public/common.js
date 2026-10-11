export const $ = (s) => document.querySelector(s);
export const money = (c) => '$' + (c / 100).toFixed(2);
export const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
export async function api(method, url, body) {
  const r = await fetch(url, { method, headers: { 'Content-Type': 'application/json' }, body: body ? JSON.stringify(body) : undefined });
  const j = await r.json();
  if (!r.ok) { toast(j.error || 'Error'); throw new Error(j.error); }
  return j;
}
export function toast(msg) {
  const t = $('.toast'); t.textContent = msg; t.classList.add('show');
  clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.remove('show'), 2200);
}
let mePromise;
export const me = () => (mePromise ||= fetch('/api/me').then((r) => (r.ok ? r.json() : (location.href = '/login', new Promise(() => {})))));
export async function nav(active) {
  document.body.insertAdjacentHTML('afterbegin', '<div class="toast"></div>');
  const u = await me();
  const links = [...(u.role !== 'checkin' ? [['snack', '🍪 Snack Sale']] : []), ...(u.role !== 'snack' ? [['checkin', '🍱 Lunch Check-in']] : []), ...(u.role === 'admin' ? [['report', '📊 Day Report'], ['settings', '⚙ Settings']] : [])];
  document.body.insertAdjacentHTML('afterbegin', `<nav>${links.map(([h, l]) => `<a href="/${h}" class="${h === active ? 'on' : ''}">${l}</a>`).join('')}
    <button id="logout" style="flex:0 0 auto;white-space:nowrap" title="Log out">${esc(u.label)} · Log out</button></nav>`);
  document.getElementById('logout').onclick = async () => { await fetch('/api/logout', { method: 'POST' }); location.href = '/login'; };
}
export const payButtons = (disabled = false) => ['cash', 'zelle', 'card']
  .map((m) => `<button class="${m}" data-m="${m}" ${disabled ? 'disabled' : ''}>${m[0].toUpperCase() + m.slice(1)}</button>`).join('');

// Tap-to-add point of sale for one menu category ('snack' or 'lunch'); logs the sale with its payment type.
export async function mountPOS(root, { category, title = 'Current sale' }) {
  const isAdmin = (await me()).role === 'admin';
  const editor = isAdmin ? `<details style="margin-top:16px"><summary>Edit menu &amp; prices</summary>
      <div class="panel row" style="margin-top:8px"><input data-r="nn" placeholder="Name"><input data-r="np" type="number" step="0.25" placeholder="Price"><button data-r="add">Add</button></div>
      <div data-r="manage"></div></details>` : '';
  root.innerHTML = `<div class="pos"><section><div class="grid" data-r="items"></div>
    ${editor}</section>
    <aside class="panel"><h3 style="margin-top:0">${title}</h3><div data-r="cart"></div>
      <div class="total" data-r="total">$0.00</div><div class="pay" data-r="pay">${payButtons(true)}<button class="vol" data-m="volunteer" disabled>🤝 Give to volunteer (free)</button></div>
      <button data-r="clear" style="width:100%;margin-top:10px">Clear</button></aside></div>`;
  const r = (n) => root.querySelector(`[data-r="${n}"]`);
  let items = [], cart = new Map();
  const price = (id) => items.find((i) => i.id == id).price_cents;
  const total = () => [...cart].reduce((s, [id, q]) => s + price(id) * q, 0);
  function drawCart() {
    r('cart').innerHTML = cart.size ? [...cart].map(([id, q]) => { const i = items.find((x) => x.id == id);
      return `<div class="cart-row"><span>${esc(i.name)}</span><button data-dec="${id}">−</button><b>${q}</b><button data-inc="${id}">+</button><span style="flex:0 0 64px;text-align:right">${money(i.price_cents * q)}</span></div>`; }).join('')
      : '<p style="color:var(--mute)">Tap items to add</p>';
    r('total').textContent = money(total());
    r('pay').querySelectorAll('button').forEach((b) => (b.disabled = !cart.size));
  }
  function render() {
    r('items').innerHTML = items.map((i) => `<button class="item" data-id="${i.id}">${i.image ? `<img src="${esc(i.image)}" alt="" loading="lazy">` : `<span class="noimg">${esc(i.name[0])}</span>`}${esc(i.name)}<small>${i.price_cents ? money(i.price_cents) : 'price not set'}</small></button>`).join('');
    if (isAdmin) r('manage').innerHTML = items.map((i) => `<div class="cart-row"><span>${esc(i.name)} — ${money(i.price_cents)}</span><button data-photo="${i.id}">Photo</button><button data-price="${i.id}">Set price</button><button data-del="${i.id}">Remove</button></div>`).join('');
    drawCart();
  }
  async function load() { items = await api('GET', '/api/snacks?category=' + category); for (const id of cart.keys()) if (!items.some((i) => i.id == id)) cart.delete(id); render(); }
  root.addEventListener('click', async (e) => {
    const t = e.target.closest('button'); if (!t) return; const d = t.dataset;
    if (d.id) { cart.set(d.id, (cart.get(d.id) || 0) + 1); drawCart(); }
    if (d.inc) { cart.set(d.inc, cart.get(d.inc) + 1); drawCart(); }
    if (d.dec) { const q = cart.get(d.dec) - 1; q ? cart.set(d.dec, q) : cart.delete(d.dec); drawCart(); }
    if (d.photo) { const f = document.createElement('input'); f.type = 'file'; f.accept = 'image/png,image/jpeg,image/webp';
      f.onchange = async () => { const file = f.files[0]; if (!file) return;
        const res = await fetch(`/api/snacks/${d.photo}/image?type=${encodeURIComponent(file.type)}`, { method: 'POST', body: file });
        if (res.ok) { toast('Picture updated'); load(); } else toast((await res.json()).error); };
      f.click(); }
    if (d.price) { const v = prompt('New price (e.g. 15.99)'); if (v !== null && v.trim() !== '') { await api('PUT', '/api/snacks/' + d.price, { price: v }); load(); } }
    if (d.del && confirm('Remove this item?')) { await api('DELETE', '/api/snacks/' + d.del); load(); }
    if (d.m) {
      const res = await api('POST', '/api/sales', { method: d.m, items: [...cart].map(([id, qty]) => ({ id: +id, qty })) });
      toast((res.volunteer ? 'Logged for volunteer ✓' : `Logged ${money(res.total_cents)} · ${d.m}`) + (res.excel ? '' : ' (Excel not updated — is the file open?)')); cart.clear(); drawCart(); root.dispatchEvent(new CustomEvent('sale', { bubbles: true })); }
    if (t === r('clear')) { cart.clear(); drawCart(); }
    if (isAdmin && t === r('add')) { await api('POST', '/api/snacks', { name: r('nn').value, price: r('np').value || 0, category }); r('nn').value = r('np').value = ''; load(); }
  });
  load();
}
