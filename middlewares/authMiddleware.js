const jwt = require('jsonwebtoken');

const autenticarJWT = (req, res, next) => {
  const authHeader = req.headers['authorization'];
  const token = authHeader?.split(' ')[1];

  if (!token) return res.status(401).json({ erro: 'Token não fornecido.' });

  jwt.verify(token, process.env.JWT_ACCESS_SECRET, (err, user) => {
    if (err) return res.status(403).json({ erro: 'Token inválido.' });

    req.user = user;
    next();
  });
};

module.exports = { autenticarJWT };
