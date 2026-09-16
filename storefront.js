'use strict';

// 完成 flower-orders-backend.gs 的網頁應用程式部署後，將網址貼在引號內。
const FLOWER_ORDER_API = '';

const $ = id => document.getElementById(id);
const form = $('orderForm');
const money = n => new Intl.NumberFormat('zh-TW').format(Number(n) || 0);

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
