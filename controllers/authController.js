const bcrypt = require('bcrypt');
const usuarioModel = require('../models/usuarioModel');
const jwt = require('jsonwebtoken');
const tokenModel = require('../models/tokenModel');

const register = async (req, res) => {
  const { nome, cpf, senha, role_id } = req.body;

  // 1. Garante que o usuário logado tem permissão (precisa do middleware antes!)
  if (req.user?.role_id !== 1) {
    return res.status(403).json({ erro: 'Acesso negado. Apenas gerentes podem registrar usuários.' });
  }

  // 2. Validação de campos
  if (!nome || !cpf || !senha || !role_id) {
    return res.status(400).json({ erro: 'Nome, CPF, senha e role_id são obrigatórios.' });
  }

  // 3. Validação do role_id informado
  if (![1, 2, 3, 4].includes(role_id)) {
    return res.status(400).json({ erro: 'role_id inválido.' });
  }

  try {
    const existente = await usuarioModel.buscarPorCpf(cpf);
    if (existente) {
      return res.status(400).json({ erro: 'CPF já cadastrado.' });
    }

    const senhaHash = await bcrypt.hash(senha, 10);
    const usuario = await usuarioModel.criarUsuario(nome, cpf, senhaHash, role_id);

    res.status(201).json({
      mensagem: 'Usuário registrado com sucesso.',
      usuarioId: usuario.insertId
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro interno no servidor.' });
  }
};


const login = async (req, res) => {
  const { cpf, senha } = req.body;
  console.log('Login attempt:', { cpf });
  console.log('Senha:', senha ? '********' : 'não fornecida');
  if (!cpf || !senha) {
    return res.status(400).json({ erro: 'CPF e senha são obrigatórios.' });
  }

  try {
    const usuario = await usuarioModel.buscarPorCpf(cpf);
    if (!usuario) {
      return res.status(401).json({ erro: 'Credenciais inválidas.' });
    }

    const senhaValida = await bcrypt.compare(senha, usuario.senha_hash);
    if (!senhaValida) {
      return res.status(401).json({ erro: 'Credenciais inválidas.' });
    }

    const payload = {
      id: usuario.id,
      nome: usuario.nome,
      role_id: usuario.role_id,
      cpf: usuario.cpf
    };

    const accessToken = jwt.sign(payload, process.env.JWT_ACCESS_SECRET, {
      expiresIn: process.env.JWT_ACCESS_EXPIRATION
    });

    const refreshToken = jwt.sign(payload, process.env.JWT_REFRESH_SECRET, {
      expiresIn: process.env.JWT_REFRESH_EXPIRATION
    });

    const agora = new Date();
    const expiracao = new Date(agora.getTime() + 1000 * 60 * 60 * 24 * 7); // 7 dias

    await tokenModel.salvarRefreshToken(usuario.id, refreshToken, expiracao);

    res.json({ accessToken, refreshToken,role_id:usuario.role_id});
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro interno no servidor.' });
  }
};

const renovarToken = async (req, res) => {
  const { refreshToken } = req.body;
  if (!refreshToken) return res.status(400).json({ erro: 'Refresh token é obrigatório.' });

  try {
    const tokenSalvo = await tokenModel.buscarToken(refreshToken);
    if (!tokenSalvo) return res.status(403).json({ erro: 'Refresh token inválido ou expirado.' });

    jwt.verify(refreshToken, process.env.JWT_REFRESH_SECRET, (err, payload) => {
      if (err) return res.status(403).json({ erro: 'Refresh token inválido.' });

      const novoAccessToken = jwt.sign(
        {
          id: payload.id,
          role_id: payload.role_id,
          cpf: payload.cpf
        },
        process.env.JWT_ACCESS_SECRET,
        { expiresIn: process.env.JWT_ACCESS_EXPIRATION }
      );

      res.json({ accessToken: novoAccessToken });
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro interno ao renovar token.' });
  }
};

const logout = async (req, res) => {
  const { refreshToken } = req.body;
  if (!refreshToken) return res.status(400).json({ erro: 'Refresh token é obrigatório.' });

  try {
    await tokenModel.removerRefreshToken(refreshToken);
    res.json({ mensagem: 'Logout realizado com sucesso.' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro interno ao realizar logout.' });
  }
};

module.exports = {
  register,
  login,
  renovarToken,
  logout
};
