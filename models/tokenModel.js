const pool = require('../utils/db')

const salvarRefreshToken = async (userId, token, expiracao) => {
    const sql = `
        INSERT INTO refresh_tokens (user_id, token, expiracao)
        VALUES (?, ?, ?)
    `
    await pool.query(sql, [userId, token, expiracao])
}
const buscarToken = async (token) => {
    const sql = `
        SELECT * FROM refresh_tokens
        WHERE token = ? AND expiracao > NOW()
        LIMIT 1
    `
    const [result] = await pool.query(sql, [token])
    return result[0] || null
}
const removerRefreshToken = async (token) => {
  const sql = `
    DELETE FROM refresh_tokens WHERE token = ?
  `
  await pool.query(sql, [token])
}

module.exports = {
  salvarRefreshToken,
  buscarToken,
  removerRefreshToken
}
