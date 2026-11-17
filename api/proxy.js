module.exports = async (req, res) => {
  const GAS_URL = process.env.GAS_URL;
  if (!GAS_URL) {
    return res.status(500).json({ error: 'GAS_URL no configurada en Vercel' });
  }

  // Configurar CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  const { method, query } = req;
  const url = new URL(GAS_URL);
  
  // Reenviar parámetros GET
  Object.keys(query).forEach(key => url.searchParams.append(key, query[key]));

  try {
    let bodyData = req.body;
    
    // Si el body viene como string (por el Content-Type text/plain), lo parseamos
    if (typeof bodyData === 'string') {
      try { bodyData = JSON.parse(bodyData); } catch(e) {}
    }

    const options = {
      method,
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      redirect: 'follow'
    };
    
    if (method === 'POST') {
      options.body = JSON.stringify(bodyData);
    }

    const response = await fetch(url.toString(), options);
    const text = await response.text();
    
    let data;
    try {
      data = JSON.parse(text);
    } catch (e) {
      return res.status(500).json({ error: 'Respuesta no válida del servidor: ' + text });
    }
    
    res.status(200).json(data);
  } catch (error) {
    res.status(500).json({ error: 'Error en Proxy: ' + error.message });
  }
};