const pool = require('../utils/db');

const buscarPorCpf = async (cpf) => {
    const [rows] = await pool.query(
        'SELECT * FROM usuarios WHERE cpf = ?',
        [cpf]
    );
    return rows[0];
};

const criarUsuario = async (nome, cpf, senhaHash, roleId) => {
  const [resultado] = await pool.query(
    'INSERT INTO usuarios (nome, cpf, senha_hash, role_id) VALUES (?, ?, ?, ?)',
    [nome, cpf, senhaHash, roleId]
  );
  return resultado;
};

module.exports = { buscarPorCpf, criarUsuario };


