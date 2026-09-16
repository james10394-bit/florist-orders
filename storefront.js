'use strict';

const FLOWER_ORDER_API = window.AMOR_CONFIG?.orderApi || '';
const PRODUCT_REFRESH_MS = window.AMOR_CONFIG?.productRefreshMs || 120000;

const $ = id => document.getElementById(id);
const form = $('orderForm');
const money = n => new Intl.NumberFormat('zh-TW').format(Number(n) || 0);

function bindFilters() {
document.querySelectorAll('[data-filter]').forEach(button => {
  button.addEventListener('click', () => {
    document.querySelectorAll('[data-filter]').forEach(x => x.classList.remove('active'));
    button.classList.add('active');
    const tag = button.dataset.filter;
    document.querySelectorAll('.product').forEach(card => {
      card.hidden = tag !== 'all' && !card.dataset.tags.split(' ').includes(tag);
    });
  });
});
}

function bindProductButtons() {
document.querySelectorAll('.product button').forEach(button => {
  button.addEventListener('click', () => {
    const card = button.closest('.product');
    $('product').value = card.dataset.name;
    $('unitPrice').value = card.dataset.price;
    $('selectedProduct').value = card.dataset.name;
    updateEstimate();
    $('order').scrollIntoView({behavior: 'smooth'});
  });
});
}

bindFilters();
bindProductButtons();

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, character => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[character]));
}

function categoryTags(category) {
  const map = {'生日':'birthday','告白／紀念':'love','告白':'love','紀念日':'love','開幕':'opening','日常':'daily'};
  return String(category || '').split(/[、,，/]/).map(x => map[x.trim()] || '').filter(Boolean).join(' ') || 'birthday love opening daily';
}

function safeImageUrl(value) {
  try {
    const url = new URL(value);
    return url.protocol === 'https:' ? url.href : '';
  } catch {
    return '';
  }
}

window.renderFlowerCatalog = payload => {
  if (!payload?.ok || !Array.isArray(payload.products) || !payload.products.length) return;
  const palettes = ['blush','rouge','sunshine','forest','violet'];
  const symbols = ['❀✿❁','✿❀✽','✺✿❀','❈✽❁','❀✾✿'];
  const cards = payload.products.map((product, index) => {
    const name = escapeHtml(product.name);
    const subtitle = escapeHtml(product.subtitle);
    const price = Number(product.price) || 0;
    const image = safeImageUrl(product.imageUrl);
    const art = image
      ? `<div class="product-art product-photo"><img src="${escapeHtml(image)}" alt="${name}" loading="lazy"></div>`
      : `<div class="product-art ${palettes[index % palettes.length]}">${symbols[index % symbols.length].split('').map(x => `<span>${x}</span>`).join('')}</div>`;
    return `<article class="product" data-tags="${categoryTags(product.category)}" data-name="${name}" data-price="${price}">${art}<div class="product-body"><div><p>${subtitle}</p><h3>${name}</h3></div><strong>${price ? 'NT$' + money(price) : '另行報價'}</strong><button>${escapeHtml(product.buttonText || '選這束')}</button></div></article>`;
  }).join('');
  cards += '<article class="product custom-card" data-tags="birthday love opening daily" data-name="花藝師客製" data-price="0"><div class="custom-inner"><span>＋</span><h3>找不到剛好的？</h3><p>告訴我們用途、預算與色系，讓花藝師為你設計。</p><button>開始客製</button></div></article>';
  $('products').innerHTML = cards;
  bindProductButtons();
  const activeFilter = document.querySelector('[data-filter].active')?.dataset.filter || 'all';
  document.querySelectorAll('.product').forEach(card => card.hidden = activeFilter !== 'all' && !card.dataset.tags.split(' ').includes(activeFilter));
};

function loadCatalog() {
  if (!FLOWER_ORDER_API) return;
  document.getElementById('catalogJsonp')?.remove();
  const script = document.createElement('script');
  script.id = 'catalogJsonp';
  script.src = FLOWER_ORDER_API + (FLOWER_ORDER_API.includes('?') ? '&' : '?') + 'action=catalog&callback=renderFlowerCatalog&_=' + Date.now();
  script.onerror = () => console.warn('商品目錄暫時無法更新，保留目前商品。');
  document.body.appendChild(script);
}

loadCatalog();
setInterval(loadCatalog, PRODUCT_REFRESH_MS);

function updateEstimate() {
  const price = Number($('unitPrice').value);
  const qty = Math.max(1, Number($('quantity').value) || 1);
  $('estimate').textContent = price ? 'NT$' + money(price * qty) : '客製報價';
}
$('quantity').addEventListener('input', updateEstimate);

const tomorrow = new Date();
tomorrow.setDate(tomorrow.getDate() + 1);
$('deliveryDate').min = [tomorrow.getFullYear(), String(tomorrow.getMonth() + 1).padStart(2, '0'), String(tomorrow.getDate()).padStart(2, '0')].join('-');

$('method').addEventListener('change', () => {
  const delivery = $('method').value === '店家配送';
  ['address', 'recipientName', 'recipientPhone'].forEach(id => $(id).required = delivery);
});

form.addEventListener('submit', event => {
  if (!form.checkValidity()) {
    event.preventDefault();
    form.reportValidity();
    return;
  }
  if (!FLOWER_ORDER_API) {
    event.preventDefault();
    $('formStatus').textContent = '店家尚未完成 Google 行事曆串接，請稍後再試或直接聯絡店家。';
    return;
  }
  if (form.website.value) {
    event.preventDefault();
    return;
  }
  const orderId = 'FL-' + Date.now().toString(36).toUpperCase() + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
  $('orderId').value = orderId;
  form.action = FLOWER_ORDER_API;
  $('submitOrder').disabled = true;
  $('formStatus').textContent = '正在送出預訂資料…';
});

window.addEventListener('message', event => {
  const data = event.data;
  if (!data || data.app !== 'flower-order') return;
  $('submitOrder').disabled = false;
  if (data.ok) {
    $('formStatus').textContent = '';
    $('successId').textContent = data.orderId || $('orderId').value;
    $('success').hidden = false;
    form.reset();
    $('product').value = '花藝師客製';
    $('unitPrice').value = '0';
    $('selectedProduct').value = '花藝師客製';
    updateEstimate();
  } else {
    $('formStatus').textContent = data.message || '送出失敗，請稍後再試。';
  }
});

$('closeSuccess').addEventListener('click', () => $('success').hidden = true);
