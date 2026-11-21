let stock, template, catalog;
function api(path, options = {}) {
  const origin = new URL(window.BAR_API_ORIGIN);
  if (origin.protocol !== 'https:') throw Error('กรุณาตรวจ URL ของ Worker ใน config.js');
  return fetch(origin.origin + path, options);
}
async function start() {
  try {
    if (!window.BAR_API_ORIGIN) { location.replace('./preview.html'); return; }
    const response = await api('/api/master');
    if (!response.ok) {
      const info = await response.json();
      throw Error(info.error || 'โหลดข้อมูลพาร์ไม่สำเร็จ');
    }
    ({stock, template, catalog} = await response.json());
    await new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = 'app.js';
      script.onload = resolve;
      script.onerror = () => reject(Error('โหลดแอปไม่สำเร็จ กรุณาเปิดลิงก์อีกครั้ง'));
      document.body.append(script);
    });
    document.getElementById('connection').hidden = true;
  } catch (error) {
    document.getElementById('connectMsg').textContent = error.message;
  }
}
start();
