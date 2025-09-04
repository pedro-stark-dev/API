module.exports = (req, res, next) => {
  const { nome,cpf, senha } = req.body || {};

  if (!nome || !cpf || !senha) return res.status(400).json({ erro: 'Nome,CPF e senha são obrigatórios.' });

  const cpfRegex = /^\d{11}$/;
  if (!cpfRegex.test(cpf)) return res.status(400).json({ erro: 'CPF inválido. Deve conter 11 dígitos numéricos.' });

  if (senha.length < 6) return res.status(400).json({ erro: 'Senha deve ter no mínimo 6 caracteres.' });

  next();
};