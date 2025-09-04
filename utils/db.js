const mysql = require('mysql2/promise')
require('dotenv').config()

const pool = mysql.createPool({
    host: process.env.DB_HOST,
    port: process.env.DB_PORT, // <-- adicione isso
    user: process.env.DB_USER,
    password: process.env.DB_PASS,
    database: process.env.DB_NAME,
});;
pool.getConnection()
.then(()=>{
    console.log("conectado ao banco de dados com sucesso")
})
.catch(err => console.error('Erro na conexão:', err))
module.exports = pool