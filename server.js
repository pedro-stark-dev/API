const express = require('express');
const { autenticarJWT } = require('./middlewares/authMiddleware');
const cors = require('cors');
const bcrypt = require('bcrypt');
const ExcelJS = require('exceljs');
require('dotenv').config();
const app = express();

// Middlewares de parsing e CORS
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Rotas de autenticação
const authRoutes = require('./routes/authRoutes');
const pool = require('./utils/db');
app.use('/auth', authRoutes);


// Rota protegida de exemplo
app.get('/perfil', autenticarJWT, (req, res) => {
  res.json({ mensagem: 'Você está autenticado', usuario: req.user });
});

app.get('/usuarios', autenticarJWT, async (req, res) => {
  try {
    const { search } = req.query;
    let sql = 'SELECT * FROM usuarios';
    const params = [];

    if (search) {
      sql += ' WHERE nome LIKE ? OR cpf LIKE ?';
      const busca = `%${search}%`;
      params.push(busca, busca);
    }

    const [usuarios] = await pool.query(sql, params);
    res.status(200).json({ usuarios });

  } catch (err) {
    console.error('Erro ao buscar usuários:', err.message);
    res.status(500).json({ erro: 'Erro interno ao buscar usuários.' });
  }
});
app.get('/usuarios/:id', autenticarJWT, async (req, res) => {
  try {
    const id = req.params.id;
    const [rows] = await pool.query('SELECT * FROM usuarios WHERE id = ?', [id]);

    if (rows.length === 0) {
      return res.status(404).json({ erro: 'Usuário não encontrado.' });
    }

    res.status(200).json(rows[0]);

  } catch (err) {
    console.error('Erro ao buscar usuário por ID:', err.message);
    res.status(500).json({ erro: 'Erro interno ao buscar usuário.' });
  }
});


app.post('/usuarios/edit', autenticarJWT, async (req, res) => {
  try {
    const { id, nome, cpf, senha, role_id } = req.body;

    if (!id || !nome || !cpf || !role_id) {
      return res.status(400).json({ erro: 'ID, nome, CPF e role_id são obrigatórios.' });
    }

    let query = '';
    let params = [];

    if (senha && senha.trim() !== '') {
      const saltRounds = 10;
      const hashSenha = await bcrypt.hash(senha, saltRounds);

      query = `
        UPDATE usuarios 
        SET nome = ?, cpf = ?, senha_hash = ?, role_id = ? 
        WHERE id = ?
      `;
      params = [nome, cpf, hashSenha, role_id, id];
    } else {
      query = `
        UPDATE usuarios 
        SET nome = ?, cpf = ?, role_id = ? 
        WHERE id = ?
      `;
      params = [nome, cpf, role_id, id];
    }

    const [result] = await pool.query(query, params);

    if (result.affectedRows === 0) {
      return res.status(404).json({ erro: 'Usuário não encontrado.' });
    }

    res.json({ mensagem: 'Usuário atualizado com sucesso.' });

  } catch (err) {
    console.error('Erro ao editar usuário:', err);
    res.status(500).json({ erro: 'Erro interno ao editar usuário.' });
  }
});



app.post('/usuarios/remove', autenticarJWT, async (req, res) => {
  try {
    const { id } = req.body;

    if (!id) {
      return res.status(400).json({ erro: 'ID é obrigatório.' });
    }

    const [result] = await pool.query('DELETE FROM usuarios WHERE id = ?', [id]);

    if (result.affectedRows === 0) {
      return res.status(404).json({ erro: 'Usuario não encontrado' });
    }

    res.json({ mensagem: 'Usuario deletado com sucesso.' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao deletar Usuario.' });
  }
});
app.get('/produtos', autenticarJWT, async (req, res) => {
  const search = req.query.search || '';
  try {
    let query = 'SELECT * FROM produtos';
    let params = [];

    if (search) {
      query += ` WHERE 
    id LIKE ? OR
    tipo LIKE ? OR
    nome LIKE ? OR
    valor LIKE ? OR
    descricao LIKE ?
    LIMIT 10`;

      const termo = `%${search}%`;
      params.push(termo, termo, termo, termo, termo);
    }

    const [produtos] = await pool.query(query, params); // <-- desestrutura para pegar só os dados
    res.json(produtos); // <-- envia apenas o array de produtos
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar produtos.' });
  }
});

app.get('/produtos_all', autenticarJWT, async (req, res) => {
  try {
    let query = 'SELECT * FROM produtos';

    const [produtos] = await pool.query(query);
    res.json(produtos);
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar produtos.' });
  }
});
// Adicionar produto (mantém igual)
app.post('/produtos/add', autenticarJWT, async (req, res) => {
  try {
    const { nome, valor, quantidade, descricao, tipo } = req.body;

    if (!nome || !valor || !quantidade || !tipo) {
      return res.status(400).json({ erro: 'Nome, valor e quantidade são obrigatórios.' });
    }
    if (tipo !== 'Material' && tipo !== 'Produto' && tipo !== 'Sobras' && tipo !== 'Bobina') {
      return res.status(400).json({ erro: 'Tipo inválido. Deve ser Material, Produto, Sobras ou Bobina.' });
    }
    const [result] = await pool.query(
      'INSERT INTO produtos (nome, valor, quantidade, descricao, tipo) VALUES (?, ?, ?, ?, ?)',
      [nome, valor, quantidade, descricao || null, tipo]
    );

    res.status(201).json({ mensagem: 'Produto adicionado com sucesso.', produtoId: result.insertId });
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao adicionar produto.' });
  }
});
app.get('/produtos/:id', autenticarJWT, async (req, res) => {
  try {
    const { id } = req.params;

    const [result] = await pool.query('SELECT * FROM produtos WHERE id = ?', [id]);

    if (result.length === 0) {
      return res.status(404).json({ erro: 'Produto não encontrado.' });
    }

    res.json(result[0]);
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar produto.' });
  }
});
// Editar produto
app.post('/produtos/edit', autenticarJWT, async (req, res) => {
  try {
    const { id, nome, valor, quantidade, descricao, tipo } = req.body;

    if (!id || !nome) {
      return res.status(400).json({ erro: 'ID, nome, valor e quantidade são obrigatórios.' });
    }
    if (parseFloat(valor) <= 0 || parseFloat(quantidade) < 0) {
      return res.status(400).json({ erro: 'Valor deve ser maior que zero e quantidade não pode ser negativa.' });
    }
    const [result] = await pool.query(
      'UPDATE produtos SET nome = ?, valor = ?, quantidade = ?, descricao = ?, tipo = ? WHERE id = ?',
      [nome, valor, quantidade, descricao, tipo || null, id]
    );

    if (result.affectedRows === 0) {
      return res.status(404).json({ erro: 'Produto não encontrado.' });
    }

    res.json({ mensagem: 'Produto atualizado com sucesso.' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao atualizar produto.' });
  }
});

// remover produto
app.post('/produtos/remove', autenticarJWT, async (req, res) => {
  const conn = await pool.getConnection();
  await conn.beginTransaction();

  try {
    const { id } = req.body;

    if (!id) {
      return res.status(400).json({ erro: 'ID é obrigatório.' });
    }

    // Deleta itens de venda vinculados ao produto
    await conn.query('DELETE FROM itens_venda WHERE produto_id = ?', [id]);

    // Deleta constituintes ligados às receitas do produto
    await conn.query(`
      DELETE c FROM constituintes c
      INNER JOIN receita r ON c.receita_id = r.id
      WHERE r.produto_id = ?
    `, [id]);

    // Deleta receitas do produto
    await conn.query('DELETE FROM receita WHERE produto_id = ?', [id]);

    // Agora deleta o produto
    const [result] = await conn.query('DELETE FROM produtos WHERE id = ?', [id]);

    if (result.affectedRows === 0) {
      await conn.rollback();
      return res.status(404).json({ erro: 'Produto não encontrado.' });
    }

    await conn.commit();
    res.json({ mensagem: 'Produto e suas dependências foram deletados com sucesso.' });

  } catch (err) {
    await conn.rollback();
    console.error(err);
    res.status(500).json({ erro: 'Erro ao deletar produto.' });
  } finally {
    conn.release();
  }
});


app.get('/vendas', autenticarJWT, async (req, res) => {
  const page = parseInt(req.query.page) || 1;
  const limit = parseInt(req.query.limit) || 10;
  const offset = (page - 1) * limit;
  const search = req.query.busca || '';

  try {
    let where = '';
    let params = [];

    if (search) {
      where = `WHERE c.nome LIKE ? OR p.nome LIKE ?`;
      const termo = `%${search}%`;
      params.push(termo, termo);
    }

    // 🔹 Contagem de vendas (DISTINCT pra não duplicar por join)
    const [[{ total }]] = await pool.query(
      `SELECT COUNT(DISTINCT v.id) AS total
       FROM vendas v
       JOIN clientes c ON v.cliente_id = c.id
       JOIN itens_venda iv ON v.id = iv.venda_id
       JOIN produtos p ON iv.produto_id = p.id
       ${where}`,
      params
    );

    const totalPages = Math.ceil(total / limit);

    // 🔹 Buscar vendas + itens
    const [rows] = await pool.query(
      `SELECT 
         v.id AS venda_id,
         v.valor_total,
         v.vendido_em,
         v.forma_pagamento,
         v.tipo_pagamento,
         v.numero_cheque,
         c.id AS cliente_id,
         c.nome AS cliente_nome,
         iv.id AS item_id,
         iv.produto_id,
         p.nome AS produto_nome,
         iv.quantidade,
         iv.valor_unitario,
         iv.valor_total AS item_total
       FROM vendas v
       JOIN clientes c ON v.cliente_id = c.id
       JOIN itens_venda iv ON v.id = iv.venda_id
       JOIN produtos p ON iv.produto_id = p.id
       ${where}
       ORDER BY v.id ASC
       LIMIT ? OFFSET ?`,
      [...params, limit, offset]
    );

    // 🔹 Agrupar por venda
    const vendas = [];
    const mapVendas = new Map();

    rows.forEach(row => {
      if (!mapVendas.has(row.venda_id)) {
        mapVendas.set(row.venda_id, {
          id: row.venda_id,
          cliente: {
            id: row.cliente_id,
            nome: row.cliente_nome
          },
          valor_total: row.valor_total,
          vendido_em: row.vendido_em,
          forma_pagamento: row.forma_pagamento,
          tipo_pagamento: row.tipo_pagamento,
          numero_cheque: row.numero_cheque,
          produtos: []
        });
        vendas.push(mapVendas.get(row.venda_id));
      }

      mapVendas.get(row.venda_id).produtos.push({
        id: row.produto_id,
        nome: row.produto_nome,
        quantidade: row.quantidade,
        valor_unitario: Number(row.valor_unitario),
        valor_total: Number(row.item_total)
      });
    });

    res.json({ data: vendas, page, total, totalPages });
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar vendas.' });
  }
});



app.get('/vendas_all', autenticarJWT, async (req, res) => {
  try {
    const [vendas] = await pool.query('SELECT * FROM vendas');
    res.json({ vendas });
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar vendas.' });
  }
});

app.get('/vendas/total', autenticarJWT, async (req, res) => {
  try {
    const [resultado] = await pool.query('SELECT SUM(valor) AS total FROM vendas');
    res.json({ total: resultado[0].total });
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao calcular o total das vendas.' });
  }
});
app.get('/vendas/historico', autenticarJWT, async (req, res) => {
  const page = parseInt(req.query.page) || 1;
  const limit = parseInt(req.query.limit) || 10;
  const offset = (page - 1) * limit;
  const search = req.query.busca || '';

  try {
    let where = '';
    let params = [];

    if (search) {
      where = `WHERE hv.cliente_nome LIKE ? OR hvi.produto_nome LIKE ?`;
      const termo = `%${search}%`;
      params.push(termo, termo);
    }

    // 🔹 Contagem total (DISTINCT para não duplicar por itens)
    const [[{ total }]] = await pool.query(
      `SELECT COUNT(DISTINCT hv.id) AS total
       FROM historico_vendas hv
       LEFT JOIN historico_venda_itens hvi ON hv.id = hvi.historico_venda_id
       ${where}`,
      params
    );

    const totalPages = Math.ceil(total / limit);

    // 🔹 Buscar vendas + itens
    const [rows] = await pool.query(
      `SELECT 
     hv.id AS venda_id,
     hv.valor_total,
     hv.vendido_em,
     hv.forma_pagamento,
     hv.tipo_pagamento,
     hv.numero_cheque,
     hv.cliente_id,
     hv.cliente_nome,
     hvi.id AS item_id,
     hvi.produto_nome,
     hvi.quantidade,
     hvi.valor_unitario,
     hvi.valor_total AS item_total
   FROM historico_vendas hv
   LEFT JOIN historico_venda_itens hvi ON hv.id = hvi.historico_venda_id
   ${where}
   ORDER BY hv.id DESC
   LIMIT ? OFFSET ?`,
      [...params, limit, offset]
    );

    // 🔹 Agrupar itens por venda
    const vendas = [];
    const mapVendas = new Map();

    rows.forEach(row => {
      if (!mapVendas.has(row.venda_id)) {
        mapVendas.set(row.venda_id, {
          id: row.venda_id,
          cliente: {
            id: row.cliente_id,
            nome: row.cliente_nome
          },
          valor_total: row.valor_total,
          vendido_em: row.vendido_em,
          forma_pagamento: row.forma_pagamento,
          tipo_pagamento: row.tipo_pagamento,
          numero_cheque: row.numero_cheque,
          produtos: []
        });
        vendas.push(mapVendas.get(row.venda_id));
      }

      if (row.item_id) {
        mapVendas.get(row.venda_id).produtos.push({
          id: row.produto_id,
          nome: row.produto_nome,
          quantidade: row.quantidade,
          valor_unitario: Number(row.valor_unitario),
          valor_total: Number(row.item_total)
        });
      }
    });

    res.json({ data: vendas, page, total, totalPages });
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar histórico de vendas.' });
  }
});


app.post('/vendas/vender', autenticarJWT, async (req, res) => {
  const conn = await pool.getConnection();
  try {
    const { cliente, produtos, forma_pagamento, tipo_pagamento, numero_cheque, cpf, cnpj } = req.body;

    if (!cliente || !Array.isArray(produtos) || produtos.length === 0) {
      return res.status(400).json({ erro: 'Cliente e lista de produtos são obrigatórios.' });
    }

    if (!forma_pagamento || !tipo_pagamento) {
      return res.status(400).json({ erro: 'Forma e tipo de pagamento são obrigatórios.' });
    }

    await conn.beginTransaction();

    // 🔹 Buscar dados básicos do cliente
    const [[clienteDados]] = await conn.query(
      'SELECT id, nome FROM clientes WHERE id = ?',
      [cliente]
    );
    if (!clienteDados) throw new Error('Cliente não encontrado.');

    // 🔹 Criar venda (cabeçalho)
    const [vendaResult] = await conn.query(
      `INSERT INTO vendas (cliente_id, valor_total, forma_pagamento, tipo_pagamento, numero_cheque) 
       VALUES (?, 0, ?, ?, ?)`,
      [cliente, forma_pagamento, tipo_pagamento, numero_cheque || null]
    );
    const vendaId = vendaResult.insertId;

    let valorTotalVenda = 0;

    for (const item of produtos) {
      const { id: produtoId, quantidade, valor_unitario } = item;

      if (!produtoId || !quantidade || !valor_unitario) {
        throw new Error('Cada produto precisa ter id, quantidade e valor_unitario.');
      }

      // 🔹 Verificar estoque
      const [rows] = await conn.query(
        'SELECT nome, quantidade FROM produtos WHERE id = ?',
        [produtoId]
      );
      if (rows.length === 0) throw new Error(`Produto com ID ${produtoId} não encontrado.`);

      const produtoNome = rows[0].nome;
      const quantidadeAtual = Number(rows[0].quantidade);
      if (quantidadeAtual < quantidade) throw new Error(`Estoque insuficiente para o produto ID ${produtoId}.`);

      // 🔹 Inserir item da venda
      await conn.query(
        `INSERT INTO itens_venda (venda_id, produto_id, quantidade, valor_unitario) 
         VALUES (?, ?, ?, ?)`,
        [vendaId, produtoId, quantidade, valor_unitario]
      );

      // 🔹 Atualizar estoque
      await conn.query(
        'UPDATE produtos SET quantidade = quantidade - ? WHERE id = ?',
        [quantidade, produtoId]
      );

      // 🔹 Somar ao valor total da venda
      valorTotalVenda += quantidade * valor_unitario;
    }

    // 🔹 Atualizar valor total da venda
    await conn.query(
      'UPDATE vendas SET valor_total = ? WHERE id = ?',
      [valorTotalVenda, vendaId]
    );

    // ==========================================
    // 🔹 SALVAR NO HISTÓRICO (snapshot)
    // ==========================================
    const [histVendaResult] = await conn.query(
      `INSERT INTO historico_vendas 
       (cliente_id, cliente_nome, cliente_cpf, cliente_cnpj, valor_total, forma_pagamento, tipo_pagamento, numero_cheque)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        clienteDados.id,
        clienteDados.nome,
        cpf || null,
        cnpj || null,
        valorTotalVenda,
        forma_pagamento,
        tipo_pagamento,
        numero_cheque || null
      ]
    );
    const historicoId = histVendaResult.insertId;

    for (const item of produtos) {
      const [[produto]] = await conn.query(
        'SELECT nome FROM produtos WHERE id = ?',
        [item.id]
      );
      await conn.query(
        `INSERT INTO historico_venda_itens 
         (historico_venda_id, produto_nome, quantidade, valor_unitario, valor_total)
         VALUES (?, ?, ?, ?, ?)`,
        [
          historicoId,
          produto ? produto.nome : '',
          item.quantidade,
          item.valor_unitario,
          item.quantidade * item.valor_unitario
        ]
      );
    }

    await conn.commit();

    res.status(201).json({
      mensagem: 'Venda registrada com sucesso.',
      vendaId,
      historicoId,
      valorTotal: valorTotalVenda,
      forma_pagamento,
      tipo_pagamento,
      numero_cheque: numero_cheque || null
    });

  } catch (err) {
    console.error(err);
    await conn.rollback();
    res.status(500).json({ erro: err.message || 'Erro ao registrar venda.' });
  } finally {
    conn.release();
  }
});








// Buscar venda por ID
app.get('/vendas/:id', autenticarJWT, async (req, res) => {
  try {
    const { id } = req.params;

    const [result] = await pool.query('SELECT * FROM historico_vendas WHERE id = ?', [id]);

    if (result.length === 0) {
      return res.status(404).json({ erro: 'Venda não encontrada.' });
    }

    res.json(result[0]);
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar venda.' });
  }
});

app.put('/vendas/:id', autenticarJWT, async (req, res) => {
  const { id } = req.params;
  const {
    cliente_id,
    valor_total,
    forma_pagamento,
    tipo_pagamento,
    numero_cheque
  } = req.body;

  if (!cliente_id || !valor_total || !tipo_pagamento) {
    return res.status(400).json({ erro: 'Campos obrigatórios faltando.' });
  }

  try {
    const [result] = await pool.query(
      `UPDATE historico_vendas
       SET cliente_id=?, 
           valor_total=?, 
           forma_pagamento=?, 
           tipo_pagamento=?, 
           numero_cheque=?
       WHERE id=?`,
      [cliente_id, valor_total, forma_pagamento || null, tipo_pagamento, numero_cheque || null, id]
    );

    if (result.affectedRows === 0) {
      return res.status(404).json({ erro: 'Venda não encontrada.' });
    }

    res.json({ mensagem: 'Venda atualizada com sucesso.' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao atualizar venda.' });
  }
});



app.post('/vendas/periodo', async (req, res) => {
  try {
    const { mesInicio, mesTermino } = req.body;

    if (!mesInicio || !mesTermino) {
      return res.status(400).json({ erro: 'mesInicio e mesTermino são obrigatórios.' });
    }

    // Construir datas início e fim do período
    const inicio = mesInicio + '-01';
    const ultimoDia = new Date(new Date(mesTermino + '-01').getFullYear(), new Date(mesTermino + '-01').getMonth() + 1, 0).getDate();
    const fim = mesTermino + '-' + String(ultimoDia).padStart(2, '0');

    // Busca as vendas
    const sqlVendas = `
      SELECT * FROM vendas
      WHERE DATE(vendido_em) BETWEEN ? AND ?
      ORDER BY vendido_em ASC
    `;
    const [vendas] = await pool.execute(sqlVendas, [inicio, fim]);

    // Soma o total das vendas (assumindo que o campo é 'valor' ou 'total')
    const sqlTotal = `
      SELECT SUM(valor) AS totalGeral FROM vendas
      WHERE DATE(vendido_em) BETWEEN ? AND ?
    `;
    const [soma] = await pool.execute(sqlTotal, [inicio, fim]);

    res.json({
      vendas,
      totalGeral: soma[0].totalGeral || 0
    });

  } catch (error) {
    console.error('Erro no /vendas/periodo:', error);
    res.status(500).json({ erro: 'Erro interno no servidor.' });
  }
});
app.post('/vendas/delete', autenticarJWT, async (req, res) => {
  try {
    const { id } = req.body;
    const [result] = await pool.query('DELETE FROM historico_vendas WHERE id = ?', [id]);
    if (result.affectedRows === 0) {
      return res.status(404).json({ erro: 'Venda nao encontrada.' });
    }
    res.json({ mensagem: 'Venda excluida com sucesso.' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao excluir venda.' });
  }
});

// Buscar máquinas
app.get('/maquinas', autenticarJWT, async (req, res) => {
  try {
    const search = req.query.search || '';
    let query = 'SELECT * FROM maquinas';
    const params = [];

    if (search) {
      query += ' WHERE nome LIKE ? LIMIT 10';
      params.push(`%${search}%`);
    }

    const [maquinas] = await pool.query(query, params);
    res.json(maquinas);
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar máquinas.' });
  }
});
app.get('/maquinas/:id', autenticarJWT, async (req, res) => {
  try {
    const { id } = req.params;

    const [result] = await pool.query('SELECT * FROM maquinas WHERE id = ?', [id]);

    if (result.length === 0) {
      return res.status(404).json({ erro: 'Máquina não encontrada.' });
    }

    res.json(result[0]);
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar maquina.' });
  }
});
// Adicionar máquina
app.post('/maquinas/add', autenticarJWT, async (req, res) => {
  try {
    const { nome } = req.body;

    if (!nome) {
      return res.status(400).json({ erro: 'O nome da máquina é obrigatório.' });
    }

    const [result] = await pool.query(
      'INSERT INTO maquinas (nome) VALUES (?)',
      [nome]
    );

    res.status(201).json({ mensagem: 'Máquina adicionada com sucesso.', maquinaId: result.insertId });
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao adicionar máquina.' });
  }
});

//editar máquina
app.post('/maquinas/edit', autenticarJWT, async (req, res) => {
  try {
    const { id, nome } = req.body;

    if (!id || !nome) {
      return res.status(400).json({ erro: 'ID e nome são obrigatórios.' });
    }

    const [result] = await pool.query(
      'UPDATE maquinas SET nome = ? WHERE id = ?',
      [nome, id]
    );

    if (result.affectedRows === 0) {
      return res.status(404).json({ erro: 'Máquina não encontrada.' });
    }

    res.json({ mensagem: 'Máquina atualizada com sucesso.' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao atualizar máquina.' });
  }
});
//remover maquina
app.post('/maquinas/remove', autenticarJWT, async (req, res) => {
  try {
    const { id } = req.body;

    if (!id) {
      return res.status(400).json({ erro: 'ID é obrigatório.' });
    }

    const [result] = await pool.query('DELETE FROM maquinas WHERE id = ?', [id]);

    if (result.affectedRows === 0) {
      return res.status(404).json({ erro: 'Máquina não encontrada.' });
    }

    res.json({ mensagem: 'Máquina deletada com sucesso.' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao deletar máquina.' });
  }
});



//buscar ficha de extrusão
app.get('/ficha_extrusao', autenticarJWT, async (req, res) => {
  try {
    const [fichas] = await pool.query('SELECT * FROM work_ficha');
    res.json(fichas);
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar fichas de extrusão.' });
  }
});


app.get('/fichas_extrusao/v1', autenticarJWT, async (req, res) => {
  const search = req.query.search || '';
  const page = parseInt(req.query.page) || 1;
  const limit = parseInt(req.query.limit) || 10;
  const offset = (page - 1) * limit;

  try {
    let whereClause = '';
    let whereParams = [];

    if (search) {
      whereClause = `
        WHERE 
          CAST(id AS CHAR) LIKE ? OR
          operador_nome LIKE ? OR
          operador_cpf LIKE ? OR
          operador_maquina LIKE ? OR
          produto LIKE ? OR
          CAST(preenchido_em AS CHAR) LIKE ? OR
          inicio LIKE ? OR
          termino LIKE ? OR
          CAST(peso AS CHAR) LIKE ? OR
          CAST(aparas AS CHAR) LIKE ? OR
          obs LIKE ?
      `;
      const termo = `%${search}%`;
      whereParams = Array(11).fill(termo);
    }

    const countQuery = `
      SELECT COUNT(*) AS total FROM work_ficha
      ${whereClause.trim()}
    `;
    const [countRows] = await pool.query(countQuery, whereParams);
    const total = countRows[0].total;
    const totalPages = Math.ceil(total / limit);

    const dataQuery = `
      SELECT * FROM work_ficha
      ${whereClause.trim()}
      ORDER BY id ASC
      LIMIT ? OFFSET ?
    `;
    const dataParams = [...whereParams, limit, offset];
    const [rows] = await pool.query(dataQuery, dataParams);

    res.json({
      data: rows,
      totalPages
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar fichas de extrusão.' });
  }
});



// rota para adicionar ficha de extrusão
app.post('/ficha_extrusao/add', autenticarJWT, async (req, res) => {
  const {
    operador_nome,
    operador_cpf,
    operador_maquina,
    inicio,
    termino,
    produto,   // nome do produto acabado
    peso,
    aparas,
    obs,
    id         // id do produto acabado
  } = req.body;

  const connection = await pool.getConnection();
  await connection.beginTransaction();

  try {
    // 1. Verifica se o produto acabado existe
    const [produtosEncontrados] = await connection.query(
      'SELECT nome FROM produtos WHERE id = ?',
      [id]
    );

    if (produtosEncontrados.length === 0) {
      await connection.rollback();
      return res.status(404).json({ erro: 'Produto não encontrado.' });
    }

    const nomeProduto = produtosEncontrados[0].nome;

    // 2. Busca a receita associada ao produto acabado
    const [receitas] = await connection.query(
      'SELECT id FROM receita WHERE produto_id = ? LIMIT 1',
      [id]
    );

    const receitaId = receitas.length > 0 ? receitas[0].id : null;

    if (!receitaId) {
      await connection.rollback();
      return res.status(400).json({ erro: 'Nenhuma receita cadastrada para este produto.' });
    }

    // 3. Busca os constituintes da receita
    const [constituintes] = await connection.query(
      'SELECT constituinte, percentual FROM constituintes WHERE receita_id = ?',
      [receitaId]
    );

    if (constituintes.length === 0) {
      await connection.rollback();
      return res.status(400).json({ erro: 'Receita não possui constituintes cadastrados.' });
    }

    // 4. Atualiza o estoque de cada insumo conforme percentual
    for (const item of constituintes) {
      const quantidadeUsada = (peso * (item.percentual / 100));

      // Verifica se insumo existe
      const [insumos] = await connection.query(
        'SELECT id, quantidade FROM produtos WHERE nome = ?',
        [item.constituinte]
      );

      if (insumos.length === 0) {
        await connection.rollback();
        return res.status(404).json({ erro: `Insumo '${item.constituinte}' não encontrado no estoque.` });
      }

      const insumo = insumos[0];

      if (insumo.quantidade < quantidadeUsada) {
        await connection.rollback();
        return res.status(400).json({
          erro: `Estoque insuficiente para o insumo '${item.constituinte}'.`
        });
      }

      // Desconta do estoque
      await connection.query(
        'UPDATE produtos SET quantidade = quantidade - ? WHERE id = ?',
        [quantidadeUsada, insumo.id]
      );
    }

    // 5. Incrementa o estoque do produto acabado
    await connection.query(
      'UPDATE produtos SET quantidade = quantidade + ? WHERE id = ?',
      [peso, id]
    );

    // 6. Registra a ficha de extrusão
    await connection.query(
      `INSERT INTO work_ficha
      (operador_nome, operador_cpf, operador_maquina, inicio, termino,
       produto, peso, aparas, obs)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        operador_nome,
        operador_cpf,
        operador_maquina,
        inicio,
        termino,
        nomeProduto,
        peso,
        aparas,
        obs
      ]
    );

    await connection.commit();
    res.status(201).json({ mensagem: 'Ficha de extrusão adicionada com sucesso!' });

  } catch (err) {
    await connection.rollback();
    console.error('Erro ao adicionar ficha:', err);
    res.status(500).json({ erro: 'Erro ao adicionar ficha', stack: err.stack, body: req.body });
  } finally {
    connection.release();
  }
});





//buscar ficha de corte
app.get('/ficha_corte', autenticarJWT, async (req, res) => {
  try {
    const [fichas] = await pool.query('SELECT * FROM work_ficha');
    res.json(fichas);
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar fichas de extrusão.' });
  }
});


app.get('/fichas_corte/v1', autenticarJWT, async (req, res) => {
  const search = req.query.search || '';
  const page = parseInt(req.query.page) || 1;
  const limit = parseInt(req.query.limit) || 10;
  const offset = (page - 1) * limit;

  try {
    let whereClause = '';
    let whereParams = [];

    if (search) {
      whereClause = `
        WHERE 
          CAST(id AS CHAR) LIKE ? OR
          operador_nome LIKE ? OR
          operador_cpf LIKE ? OR
          maquina LIKE ? OR
          turno LIKE ? OR
          sacola_dim LIKE ? OR
          CAST(total AS CHAR) LIKE ? OR
          CAST(aparas AS CHAR) LIKE ? OR
          obs LIKE ? OR
          CAST(preenchido_em AS CHAR) LIKE ?
      `;
      const termo = `%${search}%`;
      whereParams = Array(10).fill(termo);
    }

    const countQuery = `
      SELECT COUNT(*) AS total FROM corte_ficha
      ${whereClause.trim()}
    `;
    const [countRows] = await pool.query(countQuery, whereParams);
    const total = countRows[0].total;
    const totalPages = Math.ceil(total / limit);

    const dataQuery = `
  SELECT * FROM corte_ficha
  ${whereClause}
  ORDER BY id ASC
  LIMIT ? OFFSET ?
`;
    const dataParams = [...whereParams, limit, offset];
    const [rows] = await pool.query(dataQuery, dataParams);

    res.json({
      data: rows,
      totalPages
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar fichas de corte.' });
  }
});


//rota para adicionar ficha de corte
app.post('/ficha_corte/add', autenticarJWT, async (req, res) => {
  const connection = await pool.getConnection();

  try {
    const {
      operador_nome,
      operador_cpf,
      maquina,
      turno,
      sacola_dim,
      bobina,
      total,
      aparas,
      obs,
      produto_id,
      bobina_id
    } = req.body;

    if (!operador_nome || !operador_cpf || !sacola_dim || !total || !turno || !produto_id) {
      return res.status(400).json({ erro: 'Campos obrigatórios não preenchidos.' });
    }

    const totalNum = parseFloat(total);
    const aparasNum = aparas !== undefined && aparas !== '' ? parseFloat(aparas) : null;
    const produtoIdNum = parseInt(produto_id);
    const bobinaIdNum = bobina_id ? parseInt(bobina_id) : null;

    if (isNaN(totalNum) || totalNum <= 0) {
      return res.status(400).json({ erro: 'O total deve ser um número válido e maior que zero.' });
    }

    if (isNaN(produtoIdNum) || produtoIdNum <= 0) {
      return res.status(400).json({ erro: 'ID do produto inválido.' });
    }

    await connection.beginTransaction();

    // Inserir ficha de corte
    const [result] = await connection.query(
      `INSERT INTO corte_ficha 
        (operador_nome, operador_cpf, maquina, turno, sacola_dim, bobina, bobina_id, total, aparas, obs, preenchido_em)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NOW())`,
      [
        operador_nome,
        operador_cpf,
        maquina,
        turno,
        sacola_dim,
        bobina || null,
        bobinaIdNum,
        totalNum,
        aparasNum,
        obs || null
      ]
    );

    // Atualizar quantidade do produto (saida pronta)
    const [produto] = await connection.query(`SELECT id FROM produtos WHERE id = ?`, [produtoIdNum]);
    if (produto.length === 0) {
      await connection.rollback();
      return res.status(404).json({ erro: 'Produto não encontrado para atualizar estoque.' });
    }

    await connection.query(
      `UPDATE produtos SET quantidade = quantidade + ? WHERE id = ?`,
      [totalNum, produtoIdNum]
    );

    // Atualizar aparas
    if (aparasNum && aparasNum > 0) {
      const [aparasProduto] = await connection.query(
        `SELECT id FROM produtos WHERE nome = 'Aparas' LIMIT 1`
      );
      if (aparasProduto.length > 0) {
        const aparasId = aparasProduto[0].id;
        await connection.query(
          `UPDATE produtos SET quantidade = quantidade + ?, data_atualizada = NOW() WHERE id = ?`,
          [aparasNum, aparasId]
        );
      }
    }

    // Verificar se a bobina tem estoque suficiente
    if (bobinaIdNum) {
      const [bobinaProduto] = await connection.query(
        `SELECT quantidade FROM produtos WHERE id = ?`,
        [bobinaIdNum]
      );

      if (bobinaProduto.length === 0) {
        await connection.rollback();
        return res.status(404).json({ erro: 'Bobina não encontrada.' });
      }

      if (bobinaProduto[0].quantidade < totalNum) {
        await connection.rollback();
        return res.status(400).json({ erro: 'Quantidade da bobina insuficiente.' });
      }

      // Atualizar estoque da bobina
      await connection.query(
        `UPDATE produtos SET quantidade = quantidade - ? WHERE id = ?`,
        [totalNum, bobinaIdNum]
      );
    }

    // Inserir no histórico de entrada
    await connection.query(
      `INSERT INTO historico_entrada 
        (produto_id, quantidade, data_entrada, nome, operador, maquina, aparas)
       VALUES (?, ?, ?, ?, ?, ?, ?)`,
      [produtoIdNum, totalNum, new Date(), sacola_dim, operador_nome.trim(), maquina.trim(), aparasNum]
    );

    await connection.commit();

    res.status(201).json({
      mensagem: 'Ficha de corte adicionada com sucesso.',
      fichaId: result.insertId
    });

  } catch (err) {
    await connection.rollback();
    console.error('Erro ao adicionar ficha de corte:', {
      erro: err.message,
      stack: err.stack,
      body: req.body,
      timestamp: new Date().toISOString()
    });
    res.status(500).json({
      erro: 'Erro ao adicionar ficha de corte.',
      detalhes: process.env.NODE_ENV === 'development' ? err.message : null
    });
  } finally {
    connection.release();
  }
});






//rota para relatorio de extrusão por periodo
app.post('/relatorio/extrusao', autenticarJWT, async (req, res) => {
  const connection = await pool.getConnection();
  try {
    const { data_inicio, data_fim } = req.body;

    if (!data_inicio || !data_fim) {
      return res.status(400).json({ erro: 'data_inicio e data_fim são obrigatórios.' });
    }

    const query = `
      SELECT * FROM work_ficha 
      WHERE preenchido_em BETWEEN ? AND ?
    `;
    const [dados] = await connection.query(query, [data_inicio, data_fim]);

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Extrusão');

    // ... (mesmo código para montar a planilha e adicionar os dados, bordas, totais) ...

    // Estilo de borda
    const borderStyle = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    sheet.columns = [
      { header: 'data', key: 'preenchido_em', width: 15 },
      { header: 'nome', key: 'operador_nome', width: 25 },
      { header: 'cpf', key: 'operador_cpf', width: 18 },
      { header: 'maquina', key: 'operador_maquina', width: 20 },
      { header: 'inicio', key: 'inicio', width: 12 },
      { header: 'termino', key: 'termino', width: 12 },
      { header: 'produto', key: 'produto', width: 18 },
      { header: 'peso', key: 'peso', width: 10 },
      { header: 'aparas', key: 'aparas', width: 10 },
      { header: 'obs', key: 'obs', width: 30 },
    ];

    // Cabeçalho estilizado
    sheet.getRow(1).eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF00B050' }
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = borderStyle;
    });

    let totalPeso = 0;
    let totalAparas = 0;

    dados.forEach(d => {
      const row = sheet.addRow(d);
      row.eachCell(cell => {
        cell.border = borderStyle;
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      });
      totalPeso += Number(d.peso || 0);
      totalAparas += Number(d.aparas || 0);
    });

    const ultimaLinha = sheet.rowCount + 2;

    // Total Geral
    sheet.mergeCells(`F${ultimaLinha}:G${ultimaLinha}`);
    sheet.getCell(`F${ultimaLinha}`).value = 'Total Geral:';
    sheet.getCell(`F${ultimaLinha}`).font = { bold: true };
    sheet.getCell(`F${ultimaLinha}`).alignment = { horizontal: 'right' };
    sheet.getCell(`F${ultimaLinha}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF00B050' }
    };
    sheet.getCell(`H${ultimaLinha}`).value = totalPeso;
    sheet.getCell(`H${ultimaLinha}`).font = { bold: true };
    sheet.getCell(`H${ultimaLinha}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF99' }
    };

    // Total Aparas
    sheet.mergeCells(`F${ultimaLinha + 1}:G${ultimaLinha + 1}`);
    sheet.getCell(`F${ultimaLinha + 1}`).value = 'Total de Aparas:';
    sheet.getCell(`F${ultimaLinha + 1}`).font = { bold: true };
    sheet.getCell(`F${ultimaLinha + 1}`).alignment = { horizontal: 'right' };
    sheet.getCell(`F${ultimaLinha + 1}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF00B050' }
    };
    sheet.getCell(`H${ultimaLinha + 1}`).value = totalAparas;
    sheet.getCell(`H${ultimaLinha + 1}`).font = { bold: true };
    sheet.getCell(`H${ultimaLinha + 1}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF99' }
    };

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=relatorio_extrusao.xlsx');
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao gerar planilha.' });
  } finally {
    connection.release();
  }
});

///////////////////////

//rota para relatorio de extrusão
app.get('/relatorio/extrusao', autenticarJWT, async (req, res) => {
  const connection = await pool.getConnection();
  try {
    const [dados] = await connection.query('SELECT * FROM work_ficha');

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Extrusão');

    // Estilo de borda
    const borderStyle = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    // Cabeçalho
    sheet.columns = [
      { header: 'data', key: 'preenchido_em', width: 15 },
      { header: 'nome', key: 'operador_nome', width: 25 },
      { header: 'cpf', key: 'operador_cpf', width: 18 },
      { header: 'maquina', key: 'operador_maquina', width: 20 },
      { header: 'inicio', key: 'inicio', width: 12 },
      { header: 'termino', key: 'termino', width: 12 },
      { header: 'produto', key: 'produto', width: 18 },
      { header: 'peso', key: 'peso', width: 10 },
      { header: 'aparas', key: 'aparas', width: 10 },
      { header: 'obs', key: 'obs', width: 30 },
    ];

    // Estilo do cabeçalho
    sheet.getRow(1).eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF00B050' } // Verde
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }; // Branco
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = borderStyle;
    });

    // Adicionar dados e somatórios
    let totalPeso = 0;
    let totalAparas = 0;

    dados.forEach(d => {
      const row = sheet.addRow(d);
      row.eachCell(cell => {
        cell.border = borderStyle;
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      });
      totalPeso += Number(d.peso || 0);
      totalAparas += Number(d.aparas || 0);
    });

    const ultimaLinha = sheet.rowCount + 2;

    // Total Geral
    sheet.mergeCells(`F${ultimaLinha}:G${ultimaLinha}`);
    sheet.getCell(`F${ultimaLinha}`).value = 'Total Geral:';
    sheet.getCell(`F${ultimaLinha}`).font = { bold: true };
    sheet.getCell(`F${ultimaLinha}`).alignment = { horizontal: 'right' };
    sheet.getCell(`F${ultimaLinha}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF00B050' }
    };
    sheet.getCell(`H${ultimaLinha}`).value = totalPeso;
    sheet.getCell(`H${ultimaLinha}`).font = { bold: true };
    sheet.getCell(`H${ultimaLinha}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF99' } // Amarelo claro
    };

    // Total Aparas
    sheet.mergeCells(`F${ultimaLinha + 1}:G${ultimaLinha + 1}`);
    sheet.getCell(`F${ultimaLinha + 1}`).value = 'Total de Aparas:';
    sheet.getCell(`F${ultimaLinha + 1}`).font = { bold: true };
    sheet.getCell(`F${ultimaLinha + 1}`).alignment = { horizontal: 'right' };
    sheet.getCell(`F${ultimaLinha + 1}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF00B050' }
    };
    sheet.getCell(`H${ultimaLinha + 1}`).value = totalAparas;
    sheet.getCell(`H${ultimaLinha + 1}`).font = { bold: true };
    sheet.getCell(`H${ultimaLinha + 1}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF99' }
    };

    // Enviar arquivo
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=relatorio_extrusao.xlsx');
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao gerar planilha.' });
  } finally {
    connection.release();
  }
});
////////////////////////


app.post('/relatorio/corte', autenticarJWT, async (req, res) => {
  const connection = await pool.getConnection();
  try {
    const { data_inicio, data_fim } = req.body;

    if (!data_inicio || !data_fim) {
      return res.status(400).json({ erro: 'data_inicio e data_fim são obrigatórios.' });
    }

    const query = `
      SELECT * FROM corte_ficha 
      WHERE preenchido_em BETWEEN ? AND ?
    `;
    const [dados] = await connection.query(query, [data_inicio, data_fim]);

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Corte');

    const borderStyle = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    // Colunas sem 'inicio' e 'termino'
    sheet.columns = [
      { header: 'Data', key: 'preenchido_em', width: 15 },
      { header: 'Nome', key: 'operador_nome', width: 25 },
      { header: 'CPF', key: 'operador_cpf', width: 18 },
      { header: 'Máquina', key: 'maquina', width: 20 },
      { header: 'Turno', key: 'turno', width: 18 },
      { header: 'Produto', key: 'produto', width: 18 },
      { header: 'Total', key: 'total', width: 10 },
      { header: 'Aparas', key: 'aparas', width: 10 },
      { header: 'Obs', key: 'obs', width: 30 },
    ];

    sheet.getRow(1).eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF00B050' },
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = borderStyle;
    });

    let totalPeso = 0;
    let totalAparas = 0;

    dados.forEach(d => {
      const row = sheet.addRow(d);
      row.eachCell(cell => {
        cell.border = borderStyle;
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      });

      totalPeso += Number(d.total || 0);
      totalAparas += Number(d.aparas || 0);
    });

    const ultimaLinha = sheet.rowCount + 1;

    sheet.mergeCells(`E${ultimaLinha}:F${ultimaLinha}`);
    const totalGeralCell = sheet.getCell(`E${ultimaLinha}`);
    totalGeralCell.value = 'Total Geral:';
    totalGeralCell.font = { bold: true };
    totalGeralCell.alignment = { horizontal: 'right' };
    totalGeralCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF00B050' },
    };
    totalGeralCell.border = borderStyle;

    const valorTotalCell = sheet.getCell(`G${ultimaLinha}`);
    valorTotalCell.value = totalPeso;
    valorTotalCell.font = { bold: true };
    valorTotalCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF99' },
    };
    valorTotalCell.border = borderStyle;

    sheet.mergeCells(`E${ultimaLinha + 1}:F${ultimaLinha + 1}`);
    const totalAparasCell = sheet.getCell(`E${ultimaLinha + 1}`);
    totalAparasCell.value = 'Total de Aparas:';
    totalAparasCell.font = { bold: true };
    totalAparasCell.alignment = { horizontal: 'right' };
    totalAparasCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF00B050' },
    };
    totalAparasCell.border = borderStyle;

    const valorAparasCell = sheet.getCell(`G${ultimaLinha + 1}`);
    valorAparasCell.value = totalAparas;
    valorAparasCell.font = { bold: true };
    valorAparasCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF99' },
    };
    valorAparasCell.border = borderStyle;

    sheet.autoFilter = {
      from: 'A1',
      to: 'I1',
    };

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=relatorio_corte.xlsx');

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao gerar planilha.' });
  } finally {
    connection.release();
  }
});


//rota para relatorio de corte
app.get('/relatorio/corte', autenticarJWT, async (req, res) => {
  const connection = await pool.getConnection();
  try {
    const [dados] = await connection.query('SELECT * FROM corte_ficha');

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Extrusão');

    // Estilo de borda
    const borderStyle = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    // Cabeçalho
    sheet.columns = [
      { header: 'data', key: 'preenchido_em', width: 15 },
      { header: 'nome', key: 'operador_nome', width: 25 },
      { header: 'cpf', key: 'operador_cpf', width: 18 },
      { header: 'maquina', key: 'maquina', width: 20 },
      { header: 'turno', key: 'turno', width: 18 },
      { header: 'produto', key: 'sacola_dim', width: 18 },
      { header: 'total', key: 'total', width: 10 },
      { header: 'aparas', key: 'aparas', width: 10 },
      { header: 'obs', key: 'obs', width: 30 },
    ];

    // Estilo do cabeçalho
    sheet.getRow(1).eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF00B050' } // Verde
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }; // Branco
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = borderStyle;
    });

    // Adicionar dados e somatórios
    let totalPeso = 0;
    let totalAparas = 0;

    dados.forEach(d => {
      const row = sheet.addRow(d);
      row.eachCell(cell => {
        cell.border = borderStyle;
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      });
      totalPeso += Number(d.total || 0);
      totalAparas += Number(d.aparas || 0);
    });

    const ultimaLinha = sheet.rowCount + 2;

    // Total Geral (peso) na coluna G (total)
    sheet.mergeCells(`F${ultimaLinha}:F${ultimaLinha}`);
    sheet.getCell(`F${ultimaLinha}`).value = 'Total Geral:';
    sheet.getCell(`F${ultimaLinha}`).font = { bold: true };
    sheet.getCell(`F${ultimaLinha}`).alignment = { horizontal: 'right' };
    sheet.getCell(`F${ultimaLinha}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF00B050' }
    };
    sheet.getCell(`G${ultimaLinha}`).value = totalPeso;
    sheet.getCell(`G${ultimaLinha}`).font = { bold: true };
    sheet.getCell(`G${ultimaLinha}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF99' }
    };

    // Total Aparas na coluna H
    sheet.mergeCells(`F${ultimaLinha + 1}:F${ultimaLinha + 1}`);
    sheet.getCell(`F${ultimaLinha + 1}`).value = 'Total de Aparas:';
    sheet.getCell(`F${ultimaLinha + 1}`).font = { bold: true };
    sheet.getCell(`F${ultimaLinha + 1}`).alignment = { horizontal: 'right' };
    sheet.getCell(`F${ultimaLinha + 1}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF00B050' }
    };
    sheet.getCell(`G${ultimaLinha + 1}`).value = totalAparas;
    sheet.getCell(`G${ultimaLinha + 1}`).font = { bold: true };
    sheet.getCell(`G${ultimaLinha + 1}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF99' }
    };
    // Enviar arquivo
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=relatorio_corte.xlsx');
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao gerar planilha.' });
  } finally {
    connection.release();
  }
});

//relatorio de vendas geral
app.get('/relatorio/vendas', autenticarJWT, async (req, res) => {
  const connection = await pool.getConnection();
  try {
    const query = `
      SELECT 
        hv.vendido_em,
        hv.cliente_nome,
        hvi.produto_nome,
        hvi.quantidade,
        hvi.valor_unitario,
        hvi.valor_total
      FROM historico_vendas hv
      JOIN historico_venda_itens hvi ON hv.id = hvi.historico_venda_id
      ORDER BY hv.vendido_em DESC
    `;
    const [dados] = await connection.query(query);

    const ExcelJS = require('exceljs');
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Relatório de Vendas');

    const borderStyle = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    // Definir colunas
    sheet.columns = [
      { header: 'Data', key: 'vendido_em', width: 15, style: { alignment: { horizontal: 'center' } } },
      { header: 'Cliente', key: 'cliente_nome', width: 25, style: { alignment: { horizontal: 'left' } } },
      { header: 'Produto', key: 'produto_nome', width: 25, style: { alignment: { horizontal: 'left' } } },
      { header: 'Qtd', key: 'quantidade', width: 10, style: { alignment: { horizontal: 'right' } } },
      { header: 'Vlr Unit (R$)', key: 'valor_unitario', width: 15, style: { alignment: { horizontal: 'right' } } },
      { header: 'Vlr Total (R$)', key: 'valor_total', width: 15, style: { alignment: { horizontal: 'right' } } },
    ];

    // Estilo do cabeçalho
    sheet.getRow(1).eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1F4E78' },
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = borderStyle;
    });

    let totalGeral = 0;
    let totalQtd = 0;
    const totaisPorProduto = {};

    dados.forEach(d => {
      const valorTotal = Number(d.valor_total) || 0;
      const valorUnit = Number(d.valor_unitario) || 0;
      const qtd = Number(d.quantidade) || 0;

      const row = sheet.addRow({
        vendido_em: d.vendido_em,
        cliente_nome: d.cliente_nome,
        produto_nome: d.produto_nome,
        quantidade: qtd,
        valor_unitario: valorUnit,
        valor_total: valorTotal,
      });

      row.getCell('quantidade').numFmt = '#,##0';
      row.getCell('valor_unitario').numFmt = '#,##0.00';
      row.getCell('valor_total').numFmt = '#,##0.00';

      row.eachCell(cell => {
        cell.border = borderStyle;
        cell.alignment = { vertical: 'middle', horizontal: typeof cell.value === 'number' ? 'right' : 'left' };
      });

      totalQtd += qtd;
      totalGeral += valorTotal;

      const produto = d.produto_nome || 'Não informado';
      if (!totaisPorProduto[produto]) {
        totaisPorProduto[produto] = { quantidade: 0, valor: 0 };
      }
      totaisPorProduto[produto].quantidade += qtd;
      totaisPorProduto[produto].valor += valorTotal;
    });

    // Linha em branco
    sheet.addRow([]);

    const linhaTotalGeral = sheet.rowCount + 1;

    // Total Geral título
    sheet.mergeCells(`A${linhaTotalGeral}:C${linhaTotalGeral}`);
    const celTotalLabel = sheet.getCell(`A${linhaTotalGeral}`);
    celTotalLabel.value = 'Total Geral';
    celTotalLabel.font = { bold: true, size: 12 };
    celTotalLabel.alignment = { horizontal: 'right', vertical: 'middle' };
    celTotalLabel.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4BACC6' } };
    celTotalLabel.border = borderStyle;

    // Total Quantidade
    const celTotalQtd = sheet.getCell(`D${linhaTotalGeral}`);
    celTotalQtd.value = totalQtd;
    celTotalQtd.numFmt = '#,##0';
    celTotalQtd.font = { bold: true };
    celTotalQtd.alignment = { horizontal: 'right', vertical: 'middle' };
    celTotalQtd.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
    celTotalQtd.border = borderStyle;

    // Total Valor
    const celTotalValor = sheet.getCell(`F${linhaTotalGeral}`);
    celTotalValor.value = totalGeral;
    celTotalValor.numFmt = '#,##0.00';
    celTotalValor.font = { bold: true };
    celTotalValor.alignment = { horizontal: 'right', vertical: 'middle' };
    celTotalValor.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
    celTotalValor.border = borderStyle;

    // Colunas extras do total geral (unitário vazio)
    sheet.getCell(`E${linhaTotalGeral}`).border = borderStyle;

    // Espaço antes do resumo por produto
    sheet.addRow([]);

    let linhaResumo = sheet.rowCount + 1;

    // Título Resumo
    sheet.mergeCells(`A${linhaResumo}:F${linhaResumo}`);
    const celResumoTitle = sheet.getCell(`A${linhaResumo}`);
    celResumoTitle.value = 'Totais por Produto';
    celResumoTitle.font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
    celResumoTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF8064A2' } };
    celResumoTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    celResumoTitle.border = borderStyle;

    linhaResumo++;
    // Cabeçalho resumo
    sheet.getRow(linhaResumo).values = ['Produto', 'Total Qtd', 'Total Valor'];
    ['A', 'B', 'C'].forEach(col => {
      const cel = sheet.getCell(`${col}${linhaResumo}`);
      cel.font = { bold: true };
      cel.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
      cel.alignment = { horizontal: col === 'A' ? 'left' : 'right', vertical: 'middle' };
      cel.border = borderStyle;
    });

    Object.entries(totaisPorProduto).forEach(([produto, total], i) => {
      const linhaAtual = linhaResumo + 1 + i;
      sheet.getCell(`A${linhaAtual}`).value = produto;
      sheet.getCell(`B${linhaAtual}`).value = total.quantidade;
      sheet.getCell(`C${linhaAtual}`).value = total.valor;

      sheet.getCell(`B${linhaAtual}`).numFmt = '#,##0';
      sheet.getCell(`C${linhaAtual}`).numFmt = '#,##0.00';

      ['A', 'B', 'C'].forEach(col => {
        const cell = sheet.getCell(`${col}${linhaAtual}`);
        cell.border = borderStyle;
        cell.alignment = col === 'A' ? { horizontal: 'left' } : { horizontal: 'right' };
      });
    });

    // Ajustar altura
    sheet.eachRow(row => {
      row.height = 20;
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=relatorio_vendas.xlsx');
    await workbook.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao gerar planilha.' });
  } finally {
    connection.release();
  }
});




//relatorio de vendas por periodo
app.post('/relatorio/vendas/periodo', autenticarJWT, async (req, res) => {
  const connection = await pool.getConnection();
  try {
    const { data_inicio, data_fim } = req.body;

    if (!data_inicio || !data_fim) {
      return res.status(400).json({ erro: 'data_inicio e data_fim são obrigatórios.' });
    }

    const query = `
      SELECT 
        hv.vendido_em,
        hv.cliente_nome,
        hvi.produto_nome,
        hvi.quantidade,
        hvi.valor_unitario,
        hvi.valor_total
      FROM historico_vendas hv
      JOIN historico_venda_itens hvi ON hv.id = hvi.historico_venda_id
      WHERE hv.vendido_em BETWEEN ? AND ?
      ORDER BY hv.vendido_em ASC
    `;
    const [dados] = await connection.query(query, [data_inicio, data_fim]);

    const ExcelJS = require('exceljs');
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Relatório Período');

    const borderStyle = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

    // Cabeçalho
    sheet.columns = [
      { header: 'Data', key: 'vendido_em', width: 15 },
      { header: 'Cliente', key: 'cliente_nome', width: 25 },
      { header: 'Produto', key: 'produto_nome', width: 25 },
      { header: 'Quantidade', key: 'quantidade', width: 12 },
      { header: 'Valor Unitário (R$)', key: 'valor_unitario', width: 18 },
      { header: 'Valor Total (R$)', key: 'valor_total', width: 18 },
    ];

    sheet.getRow(1).eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00B050' } };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = borderStyle;
    });

    // Dados + totais
    let totalGeral = 0;
    const totaisPorProduto = {};

    dados.forEach(d => {
      const row = sheet.addRow(d);
      row.eachCell(cell => {
        cell.border = borderStyle;
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      });

      totalGeral += Number(d.valor_total || 0);

      const produto = d.produto_nome || 'Não informado';
      if (!totaisPorProduto[produto]) {
        totaisPorProduto[produto] = { quantidade: 0, valor: 0 };
      }
      totaisPorProduto[produto].quantidade += Number(d.quantidade || 0);
      totaisPorProduto[produto].valor += Number(d.valor_total || 0);
    });

    const ultimaLinha = sheet.rowCount + 2;

    // Total Geral
    sheet.mergeCells(`D${ultimaLinha}:E${ultimaLinha}`);
    sheet.getCell(`D${ultimaLinha}`).value = 'Total Geral (R$):';
    sheet.getCell(`D${ultimaLinha}`).font = { bold: true };
    sheet.getCell(`D${ultimaLinha}`).alignment = { horizontal: 'right' };
    sheet.getCell(`D${ultimaLinha}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00B050' } };

    sheet.getCell(`F${ultimaLinha}`).value = totalGeral;
    sheet.getCell(`F${ultimaLinha}`).numFmt = 'R$ #,##0.00';
    sheet.getCell(`F${ultimaLinha}`).font = { bold: true };
    sheet.getCell(`F${ultimaLinha}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF99' } };

    // Totais por Produto
    let linhaResumo = ultimaLinha + 3;
    sheet.mergeCells(`D${linhaResumo}:F${linhaResumo}`);
    sheet.getCell(`D${linhaResumo}`).value = 'Totais por Produto';
    sheet.getCell(`D${linhaResumo}`).font = { bold: true };
    sheet.getCell(`D${linhaResumo}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
    sheet.getCell(`D${linhaResumo}`).alignment = { horizontal: 'center' };

    linhaResumo++;
    sheet.getCell(`D${linhaResumo}`).value = 'Produto';
    sheet.getCell(`E${linhaResumo}`).value = 'Quantidade Total';
    sheet.getCell(`F${linhaResumo}`).value = 'Valor Total (R$)';
    ['D', 'E', 'F'].forEach(col => {
      const cell = sheet.getCell(`${col}${linhaResumo}`);
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' } };
      cell.alignment = { horizontal: 'center' };
      cell.border = borderStyle;
    });

    Object.entries(totaisPorProduto).forEach(([produto, total], i) => {
      const linhaAtual = linhaResumo + 1 + i;
      sheet.getCell(`D${linhaAtual}`).value = produto;
      sheet.getCell(`E${linhaAtual}`).value = total.quantidade;
      sheet.getCell(`F${linhaAtual}`).value = total.valor;
      sheet.getCell(`F${linhaAtual}`).numFmt = 'R$ #,##0.00';

      ['D', 'E', 'F'].forEach(col => {
        const cell = sheet.getCell(`${col}${linhaAtual}`);
        cell.border = borderStyle;
        cell.alignment = { horizontal: 'left' };
      });
    });

    // Exporta
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=relatorio_vendas_periodo.xlsx');
    await workbook.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao gerar planilha.' });
  } finally {
    connection.release();
  }
});




//relatorio de produtos
app.post('/relatorio/produtos', autenticarJWT, async (req, res) => {
  const connection = await pool.getConnection();
  try {

    const query = `
      SELECT * FROM produtos
    `;
    const [dados] = await connection
    const ExcelJS = require('exceljs');
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Extrusão');

    const borderStyle = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    sheet.columns = [
      { header: 'Data', key: 'entrou_em', width: 15 },
      { header: 'Cliente', key: 'nome', width: 18 },
      { header: 'Tipo', key: 'tipo', width: 18 },
      { header: 'Valor', key: 'valor', width: 20 },
      { header: 'Peso', key: 'quantidade', width: 12 },
      { header: 'Descrição', key: 'descricao', width: 18 },
    ];

    sheet.getRow(1).eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF00B050' }
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = borderStyle;
    });

    let totalPeso = 0;
    const totaisPorProduto = {};

    dados.forEach(d => {
      const row = sheet.addRow(d);
      row.eachCell(cell => {
        cell.border = borderStyle;
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      });

      totalPeso += Number(d.peso || 0);

      const produto = d.produto || 'Não informado';
      if (!totaisPorProduto[produto]) {
        totaisPorProduto[produto] = { peso: 0, valor: 0 };
      }
      totaisPorProduto[produto].peso += Number(d.peso || 0);
      totaisPorProduto[produto].valor += Number(d.valor || 0);
    });

    const ultimaLinha = sheet.rowCount + 2;

    // Total Geral
    sheet.mergeCells(`F${ultimaLinha}:G${ultimaLinha}`);
    sheet.getCell(`F${ultimaLinha}`).value = 'Total Geral:';
    sheet.getCell(`F${ultimaLinha}`).font = { bold: true };
    sheet.getCell(`F${ultimaLinha}`).alignment = { horizontal: 'right' };
    sheet.getCell(`F${ultimaLinha}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF00B050' }
    };
    sheet.getCell(`H${ultimaLinha}`).value = totalPeso;
    sheet.getCell(`H${ultimaLinha}`).font = { bold: true };
    sheet.getCell(`H${ultimaLinha}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF99' }
    };

    // Totais por Produto
    let linhaResumo = ultimaLinha + 3;

    sheet.mergeCells(`F${linhaResumo}:H${linhaResumo}`);
    sheet.getCell(`F${linhaResumo}`).value = 'Totais por Produto';
    sheet.getCell(`F${linhaResumo}`).font = { bold: true };
    sheet.getCell(`F${linhaResumo}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF4472C4' }
    };
    sheet.getCell(`F${linhaResumo}`).alignment = { horizontal: 'center' };

    linhaResumo++;
    sheet.getCell(`F${linhaResumo}`).value = 'Produto';
    sheet.getCell(`G${linhaResumo}`).value = 'Total Peso';
    sheet.getCell(`H${linhaResumo}`).value = 'Total Valor';
    ['F', 'G', 'H'].forEach(col => {
      const cell = sheet.getCell(`${col}${linhaResumo}`);
      cell.font = { bold: true };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFBDD7EE' }
      };
      cell.alignment = { horizontal: 'center' };
      cell.border = borderStyle;
    });

    Object.entries(totaisPorProduto).forEach(([produto, total], i) => {
      const linhaAtual = linhaResumo + 1 + i;
      sheet.getCell(`F${linhaAtual}`).value = produto;
      sheet.getCell(`G${linhaAtual}`).value = total.peso;
      sheet.getCell(`H${linhaAtual}`).value = total.valor;

      ['F', 'G', 'H'].forEach(col => {
        const cell = sheet.getCell(`${col}${linhaAtual}`);
        cell.border = borderStyle;
        cell.alignment = { horizontal: 'left' };
      });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=relatorio_extrusao.xlsx');
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao gerar planilha.' });
  } finally {
    connection.release();
  }
});


//relatorio de produtos
app.get('/relatorio/produtos', autenticarJWT, async (req, res) => {
  const connection = await pool.getConnection();
  try {
    const query = `SELECT * FROM produtos`;
    const [dados] = await connection.query(query);

    const ExcelJS = require('exceljs');
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Produtos');

    const borderStyle = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    // Configurar colunas
    sheet.columns = [
      { header: 'Data Entrada', key: 'entrou_em', width: 18 },
      { header: 'Atualizado em', key: 'data_atualizada', width: 20 },
      { header: 'Nome', key: 'nome', width: 20 },
      { header: 'Tipo', key: 'tipo', width: 15 },
      { header: 'Valor (R$)', key: 'valor', width: 15 },
      { header: 'Quantidade', key: 'quantidade', width: 15 },
      { header: 'Descrição', key: 'descricao', width: 25 },
    ];

    // Estilizar cabeçalho
    sheet.getRow(1).eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1F4E78' }
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = borderStyle;
    });

    let totalQuantidade = 0;
    const totaisPorProduto = {};

    dados.forEach(prod => {
      const quantidade = Number(prod.quantidade) || 0;
      const valor = Number(prod.valor) || 0;

      const row = sheet.addRow({
        entrou_em: prod.entrou_em,
        data_atualizada: prod.data_atualizada,
        nome: prod.nome,
        tipo: prod.tipo,
        valor: valor,
        quantidade: quantidade,
        descricao: prod.descricao,
      });

      row.eachCell(cell => {
        cell.border = borderStyle;
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      });

      row.getCell('valor').numFmt = '#,##0.00';
      totalQuantidade += quantidade;

      const chave = `${prod.nome} (${prod.tipo})`;
      if (!totaisPorProduto[chave]) {
        totaisPorProduto[chave] = { quantidade: 0, valor: 0 };
      }
      totaisPorProduto[chave].quantidade += quantidade;
      totaisPorProduto[chave].valor += valor;
    });

    // Linha separadora
    sheet.addRow([]);

    const linhaTotal = sheet.rowCount + 1;

    sheet.mergeCells(`A${linhaTotal}:E${linhaTotal}`);
    const celTotalLabel = sheet.getCell(`A${linhaTotal}`);
    celTotalLabel.value = 'Total Geral de Quantidade';
    celTotalLabel.font = { bold: true };
    celTotalLabel.alignment = { horizontal: 'right' };
    celTotalLabel.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4BACC6' } };
    celTotalLabel.border = borderStyle;

    const celTotal = sheet.getCell(`F${linhaTotal}`);
    celTotal.value = totalQuantidade;
    celTotal.font = { bold: true };
    celTotal.alignment = { horizontal: 'right' };
    celTotal.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
    celTotal.border = borderStyle;

    // Espaço
    sheet.addRow([]);
    let linhaResumo = sheet.rowCount + 1;

    // Cabeçalho Totais por Produto
    sheet.mergeCells(`A${linhaResumo}:G${linhaResumo}`);
    const celResumoTitle = sheet.getCell(`A${linhaResumo}`);
    celResumoTitle.value = 'Totais por Produto';
    celResumoTitle.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    celResumoTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: 'FF8064A2' };
    celResumoTitle.alignment = { horizontal: 'center' };
    celResumoTitle.border = borderStyle;

    linhaResumo++;
    sheet.getRow(linhaResumo).values = ['Produto', 'Total Quantidade', 'Total Valor'];
    ['A', 'B', 'C'].forEach(col => {
      const cell = sheet.getCell(`${col}${linhaResumo}`);
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
      cell.alignment = { horizontal: col === 'A' ? 'left' : 'right' };
      cell.border = borderStyle;
    });

    Object.entries(totaisPorProduto).forEach(([produto, total], i) => {
      const linha = linhaResumo + 1 + i;
      sheet.getCell(`A${linha}`).value = produto;
      sheet.getCell(`B${linha}`).value = total.quantidade;
      sheet.getCell(`C${linha}`).value = total.valor;
      sheet.getCell(`C${linha}`).numFmt = '#,##0.00';

      ['A', 'B', 'C'].forEach(col => {
        const cell = sheet.getCell(`${col}${linha}`);
        cell.border = borderStyle;
        cell.alignment = { horizontal: col === 'A' ? 'left' : 'right' };
      });
    });

    // Altura padrão
    sheet.eachRow(row => {
      row.height = 20;
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=relatorio_produtos.xlsx');
    await workbook.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao gerar planilha.' });
  } finally {
    connection.release();
  }
});
app.get('/historico/entrada', autenticarJWT, async (req, res) => {
  const connection = await pool.getConnection();
  try {
    const { mes } = req.query;

    let query;
    let params = [];

    if (mes) {
      // Filtrar por mês e agrupar por dia
      query = `
        SELECT 
          DAY(data_entrada) AS dia,
          SUM(quantidade) AS total
        FROM historico_entrada
        WHERE MONTH(data_entrada) = ?
        GROUP BY dia
        ORDER BY dia
      `;
      params = [mes];
    } else {
      // Sem filtro = retorna agrupado por mês
      query = `
        SELECT 
          MONTH(data_entrada) - 1 AS mes, 
          SUM(quantidade) AS total
        FROM historico_entrada
        GROUP BY mes
        ORDER BY mes
      `;
    }

    const [dados] = await connection.query(query, params);
    res.json(dados);
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar histórico de entradas.' });
  } finally {
    connection.release();
  }
});
app.get('/historico/entrada/v1', autenticarJWT, async (req, res) => {
  const search = req.query.search || '';
  const page = parseInt(req.query.page) || 1;
  const limit = parseInt(req.query.limit) || 10;
  const offset = (page - 1) * limit;

  try {
    let whereClause = '';
    let whereParams = [];

    if (search) {
      whereClause = `
        WHERE 
          CAST(id AS CHAR) LIKE ? OR
          CAST(produto_id AS CHAR) LIKE ? OR
          CAST(quantidade AS CHAR) LIKE ? OR
          operador LIKE ? OR
          nome LIKE ? OR
          maquina LIKE ? OR
          CAST(data_entrada AS CHAR) LIKE ?
      `;
      const termo = `%${search}%`;
      whereParams = [termo, termo, termo, termo, termo, termo, termo];
    }

    const countQuery = `
      SELECT COUNT(*) AS total FROM historico_entrada
      ${whereClause.trim()}
    `;
    const [countRows] = await pool.query(countQuery, whereParams);
    const total = countRows[0].total;
    const totalPages = Math.ceil(total / limit);

    const dataQuery = `
      SELECT * FROM historico_entrada
      ${whereClause.trim()}
      ORDER BY id ASC
      LIMIT ? OFFSET ?
    `;
    const dataParams = [...whereParams, limit, offset];
    const [rows] = await pool.query(dataQuery, dataParams);

    res.json({
      data: rows,
      totalPages
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar produtos no histórico.' });
  }
});






app.get('/historico/entrada/v2', autenticarJWT, async (req, res) => {
  try {
    let query = 'SELECT * FROM historico_entrada';

    const [historico] = await pool.query(query);
    res.json(historico);
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar histórico.' });
  }

});
app.get('/historico/entrada/relatorio', autenticarJWT, async (req, res) => {
  const connection = await pool.getConnection();
  const search = req.query.search || '';

  try {
    let whereClause = '';
    let params = [];

    if (search) {
      whereClause = `
    WHERE 
      CAST(id AS CHAR) LIKE ? OR
      CAST(produto_id AS CHAR) LIKE ? OR
      CAST(quantidade AS CHAR) LIKE ? OR
      operador LIKE ? OR
      nome LIKE ? OR
      maquina LIKE ? OR
      CAST(data_entrada AS CHAR) LIKE ?
  `;
      const termo = `%${search}%`;
      params = [termo, termo, termo, termo, termo, termo, termo];
    }


    const query = `SELECT * FROM historico_entrada ${whereClause} ORDER BY data_entrada DESC`;
    const [dados] = await connection.query(query, params);

    const ExcelJS = require('exceljs');
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Histórico de Entrada');

    const borderStyle = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    sheet.columns = [
      { header: 'Data', key: 'data_entrada', width: 15 },
      { header: 'Produto ID', key: 'produto_id', width: 15 },
      { header: 'Quantidade', key: 'quantidade', width: 12 },
      { header: 'Aparas', key: 'aparas', width: 15 },
      { header: 'Operador', key: 'operador', width: 30 },
      { header: 'Maquina', key: 'maquina', width: 18 },
    ];

    sheet.getRow(1).eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF00B050' } // Verde
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }; // Branco
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = borderStyle;
    });

    dados.forEach(d => {
      const row = sheet.addRow(d);
      row.eachCell(cell => {
        cell.border = borderStyle;
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=relatorio_historico_entrada.xlsx');
    await workbook.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao gerar planilha.' });
  } finally {
    connection.release();
  }
});




app.get('/fichas_extrusao/v1/relatorio', autenticarJWT, async (req, res) => {
  const connection = await pool.getConnection();
  const search = req.query.search || '';

  try {
    let whereClause = '';
    let params = [];

    if (search) {
      whereClause = `
        WHERE 
          CAST(id AS CHAR) LIKE ? OR
          operador_nome LIKE ? OR
          operador_cpf LIKE ? OR
          operador_maquina LIKE ? OR
          produto LIKE ? OR
          inicio LIKE ? OR
          termino LIKE ? OR
          CAST(peso AS CHAR) LIKE ? OR
          CAST(aparas AS CHAR) LIKE ? OR
          obs LIKE ? OR
          CAST(preenchido_em AS CHAR) LIKE ?
      `;
      const termo = `%${search}%`;
      params = Array(11).fill(termo);
    }

    const query = `SELECT * FROM work_ficha ${whereClause} ORDER BY preenchido_em DESC`;
    const [dados] = await connection.query(query, params);

    const ExcelJS = require('exceljs');
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Fichas de Extrusão');

    const borderStyle = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    sheet.columns = [
      { header: 'ID', key: 'id', width: 8 },
      { header: 'Operador', key: 'operador_nome', width: 25 },
      { header: 'CPF', key: 'operador_cpf', width: 18 },
      { header: 'Máquina', key: 'operador_maquina', width: 20 },
      { header: 'Produto', key: 'produto', width: 25 },
      { header: 'Peso (kg)', key: 'peso', width: 12 },
      { header: 'Aparas (kg)', key: 'aparas', width: 14 },
      { header: 'Início', key: 'inicio', width: 20 },
      { header: 'Término', key: 'termino', width: 20 },
      { header: 'Observações', key: 'obs', width: 30 },
      { header: 'Preenchido em', key: 'preenchido_em', width: 20 },
    ];

    // Cabeçalho
    sheet.getRow(1).eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF00B050' } // Verde
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }; // Branco
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = borderStyle;
    });

    let totalPeso = 0;
    let totalAparas = 0;

    // Linhas
    dados.forEach(d => {
      totalPeso += Number(d.peso) || 0;
      totalAparas += Number(d.aparas) || 0;

      const row = sheet.addRow(d);
      row.eachCell(cell => {
        cell.border = borderStyle;
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      });
    });

    // Linha de totais
    const totalRow = sheet.addRow({
      produto: 'TOTAIS:',
      peso: totalPeso,
      aparas: totalAparas,
    });

    totalRow.eachCell((cell, colNumber) => {
      cell.border = borderStyle;
      cell.font = { bold: true };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' }, // Amarelo
      };
      if (colNumber === 5 || colNumber === 6) {
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
      } else {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      }
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=relatorio_fichas_extrusao.xlsx');
    await workbook.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao gerar planilha.' });
  } finally {
    connection.release();
  }
});


app.get('/fichas_corte/v1/relatorio', autenticarJWT, async (req, res) => {
  const connection = await pool.getConnection();
  const search = req.query.search || '';

  try {
    let whereClause = '';
    let params = [];

    if (search) {
      whereClause = `
        WHERE 
          CAST(id AS CHAR) LIKE ? OR
          operador_nome LIKE ? OR
          operador_cpf LIKE ? OR
          maquina LIKE ? OR
          turno LIKE ? OR
          sacola_dim LIKE ? OR
          CAST(total AS CHAR) LIKE ? OR
          CAST(aparas AS CHAR) LIKE ? OR
          obs LIKE ? OR
          CAST(preenchido_em AS CHAR) LIKE ?
      `;
      const termo = `%${search}%`;
      params = Array(10).fill(termo);
    }

    const query = `SELECT * FROM corte_ficha ${whereClause} ORDER BY preenchido_em DESC`;
    const [dados] = await connection.query(query, params);

    const ExcelJS = require('exceljs');
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Fichas de Corte');

    const borderStyle = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };

    sheet.columns = [
      { header: 'ID', key: 'id', width: 8 },
      { header: 'Operador', key: 'operador_nome', width: 25 },
      { header: 'CPF', key: 'operador_cpf', width: 18 },
      { header: 'Máquina', key: 'maquina', width: 20 },
      { header: 'Turno', key: 'turno', width: 15 },
      { header: 'Dim. Sacola', key: 'sacola_dim', width: 20 },
      { header: 'Total (kg)', key: 'total', width: 14 },
      { header: 'Aparas (kg)', key: 'aparas', width: 14 },
      { header: 'Observações', key: 'obs', width: 30 },
      { header: 'Preenchido em', key: 'preenchido_em', width: 22 },
    ];

    // Cabeçalho
    sheet.getRow(1).eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF0070C0' } // Azul
      };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }; // Branco
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = borderStyle;
    });

    let totalGeral = 0;
    let aparasGeral = 0;

    dados.forEach(d => {
      totalGeral += Number(d.total) || 0;
      aparasGeral += Number(d.aparas) || 0;

      const row = sheet.addRow({
        ...d,
        preenchido_em: new Date(d.preenchido_em).toLocaleString('pt-BR'),
      });

      row.eachCell(cell => {
        cell.border = borderStyle;
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      });
    });

    // Linha de totais
    const totalRow = sheet.addRow({
      sacola_dim: 'TOTAIS:',
      total: totalGeral,
      aparas: aparasGeral,
    });

    totalRow.eachCell((cell, colNumber) => {
      cell.border = borderStyle;
      cell.font = { bold: true };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' }, // Amarelo
      };
      if (colNumber === 7 || colNumber === 8) {
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
      } else {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      }
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=relatorio_fichas_corte.xlsx');
    await workbook.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao gerar planilha.' });
  } finally {
    connection.release();
  }
});




app.get('/receita', async (req, res) => {
  let sql = "SELECT * FROM receita"
  try {
    const [rows] = await pool.query(sql);
    res.status(200).json(rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar receita.' });
  }
});
app.get('/receita/constituintes', async (req, res) => {
  let sql = "SELECT * FROM constituintes"
  try {
    const [rows] = await pool.query(sql);
    res.status(200).json(rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar receita.' });
  }
});



app.post('/receita/add', autenticarJWT, async (req, res) => {
  const { constituintes } = req.body;

  if (!Array.isArray(constituintes) || constituintes.length === 0) {
    return res.status(400).json({ erro: 'Constituintes são obrigatórios.' });
  }

  const conn = await pool.getConnection();
  await conn.beginTransaction();

  try {
    const tipo = 'bobina'; // receita para todos os produtos do tipo bobina

    // Busca todos os produtos do tipo "bobina"
    let [produtos] = await conn.query(
      'SELECT id FROM produtos WHERE tipo = ?',
      [tipo]
    );

    // Se não existir nenhum produto do tipo bobina, cria um genérico
    if (produtos.length === 0) {
      const [novoProduto] = await conn.query(
        'INSERT INTO produtos (nome, tipo) VALUES (?, ?)',
        ['Bobina', tipo]
      );
      produtos = [{ id: novoProduto.insertId }];
    }

    // Valida soma de percentuais
    const totalPercentual = constituintes.reduce((sum, c) => sum + Number(c.percentual), 0);
    if (totalPercentual !== 100) {
      throw new Error('A soma dos percentuais deve ser 100%.');
    }

    // Aplica a mesma receita para todos os produtos do tipo bobina
    for (const produto of produtos) {
      let produto_id = produto.id;

      // Verifica se já existe receita para o produto
      const [receitas] = await conn.query(
        'SELECT id FROM receita WHERE produto_id = ? LIMIT 1',
        [produto_id]
      );

      let receita_id;
      if (receitas.length === 0) {
        // Cria nova receita
        const [rec] = await conn.query(
          'INSERT INTO receita (produto_id) VALUES (?)',
          [produto_id]
        );
        receita_id = rec.insertId;
      } else {
        receita_id = receitas[0].id;
        // Remove constituintes antigos
        await conn.query('DELETE FROM constituintes WHERE receita_id = ?', [receita_id]);
      }

      // Insere constituintes
      for (const item of constituintes) {
        if (!item.nome || isNaN(item.percentual) || item.percentual <= 0) {
          throw new Error('Constituinte inválido.');
        }
        await conn.query(
          'INSERT INTO constituintes (receita_id, constituinte, percentual) VALUES (?, ?, ?)',
          [receita_id, item.nome, item.percentual]
        );
      }
    }

    await conn.commit();
    res.status(201).json({ mensagem: 'Receita aplicada a todos os produtos do tipo bobina com sucesso.' });

  } catch (err) {
    await conn.rollback();
    console.error(err);
    res.status(500).json({ erro: err.message || 'Erro ao cadastrar/atualizar receita.' });
  } finally {
    conn.release();
  }
});








app.get('/clientes/all', autenticarJWT, async (req, res) => {
  const connection = await pool.getConnection();
  try {
    const query = 'SELECT * FROM clientes ORDER BY nome';
    const [clientes] = await connection.query(query);
    res.json(clientes);
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: 'Erro ao buscar clientes.' });
  } finally {
    connection.release();
  }
});
//cadastro de clientes
app.post('/clientes/add', autenticarJWT, async (req, res) => {
  const { nome, tipo, cpf, cnpj, telefone, email, endereco } = req.body;

  // Validação básica
  if (!nome || !tipo) {
    return res.status(400).json({ erro: 'Nome e tipo são obrigatórios.' });
  }

  // Validação extra: CPF só para FISICA, CNPJ só para JURIDICA
  if (tipo === 'FISICA' && !cpf) {
    return res.status(400).json({ erro: 'CPF é obrigatório para pessoa física.' });
  }
  if (tipo === 'JURIDICA' && !cnpj) {
    return res.status(400).json({ erro: 'CNPJ é obrigatório para pessoa jurídica.' });
  }

  const conn = await pool.getConnection();
  try {
    const [result] = await conn.query(
      `INSERT INTO clientes 
        (nome, tipo, cpf, cnpj, telefone, email, endereco) 
       VALUES (?, ?, ?, ?, ?, ?, ?)`,
      [nome, tipo, cpf || null, cnpj || null, telefone || null, email || null, endereco || null]
    );

    res.status(201).json({
      mensagem: 'Cliente adicionado com sucesso.',
      id: result.insertId
    });

  } catch (err) {
    console.error(err);

    // Tratamento de erro mais específico
    if (err.code === 'ER_DUP_ENTRY') {
      return res.status(400).json({ erro: 'CPF, CNPJ ou e-mail já cadastrado.' });
    }

    res.status(500).json({ erro: 'Erro ao adicionar cliente.' });
  } finally {
    conn.release();
  }
});

// iniciar aplicação
app.listen(3000, () => console.log('Servidor rodando na porta 3000'));