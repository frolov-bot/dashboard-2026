const https = require('https');
const http = require('http');

const PUBLIC_KEY = 'https://disk.360.yandex.ru/i/pbXk6jtc2OeG5Q';

function fetchUrl(url, redirectCount = 0) {
  return new Promise((resolve, reject) => {
    if (redirectCount > 5) return reject(new Error('Слишком много редиректов'));

    const lib = url.startsWith('https') ? https : http;
    const req = lib.get(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; NetlifyFunction/1.0)',
      }
    }, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return resolve(fetchUrl(res.headers.location, redirectCount + 1));
      }

      if (res.statusCode !== 200) {
        return reject(new Error(`HTTP ${res.statusCode}`));
      }

      const chunks = [];
      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => resolve({
        buffer: Buffer.concat(chunks),
        contentType: res.headers['content-type'] || 'application/octet-stream'
      }));
      res.on('error', reject);
    });

    req.on('error', reject);
    req.setTimeout(30000, () => { req.destroy(); reject(new Error('Таймаут')); });
  });
}

exports.handler = async function(event, context) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers, body: '' };
  }

  try {
    const apiUrl = `https://cloud-api.yandex.net/v1/disk/public/resources/download?public_key=${encodeURIComponent(PUBLIC_KEY)}`;
    const apiResult = await fetchUrl(apiUrl);
    const apiData = JSON.parse(apiResult.buffer.toString('utf8'));

    if (!apiData.href) {
      throw new Error('Яндекс Диск не вернул ссылку на скачивание');
    }

    const fileResult = await fetchUrl(apiData.href);

    return {
      statusCode: 200,
      headers: {
        ...headers,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Length': fileResult.buffer.length.toString(),
        'Cache-Control': 'no-cache',
      },
      body: fileResult.buffer.toString('base64'),
      isBase64Encoded: true,
    };

  } catch (err) {
    console.error('Ошибка прокси:', err.message);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
