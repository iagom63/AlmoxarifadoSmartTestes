const express   = require('express');
const mongoose  = require('mongoose');
const socketIO  = require('socket.io');
const http      = require('http');
const path      = require('path');
const fs        = require('fs');
const ExcelJS   = require('exceljs');

const app    = express();
const server = http.createServer(app);
const io     = socketIO(server);
const PORT   = process.env.PORT || 3000;

// ─── Middlewares ───────────────────────────────────────────────────────────────
app.use(express.static('public'));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// ─── Conexão MongoDB ────────────────────────────────────────────────────────────
mongoose.connect(
  'mongodb+srv://ambientetestes63:JFbkt6dFlkgyJV9j@teste.tmub42r.mongodb.net/Teste?retryWrites=true&w=majority&appName=Teste',
  { useNewUrlParser: true, useUnifiedTopology: true }
);
mongoose.connection.on('error', err => console.error('❌ MongoDB error:', err));
mongoose.connection.once('open', () => console.log('✅ MongoDB conectado'));
// ─── Schema & Model ─────────────────────────────────────────────────────────────
const saidaSchema = new mongoose.Schema({
  nome: String,
  descricao: String,
  quantidade: String,
  unidade: String,
  tipo: String,
  data: String,
  observacao: String,
  responsavel: String,
  delivered: Boolean,
  visivel:   { type: Boolean, default: true },
  timestamp: Number,
  descontinuado: { type: Boolean, default: false }   // ← novo campo
});
const Saida = mongoose.model('Saida', saidaSchema);

// ─── Helpers ───────────────────────────────────────────────────────────────────
const getCurrentDate = () => new Date().toISOString().split('T')[0];

// ─── Rotas ──────────────────────────────────────────────────────────────────────

// Listagem exclui removidos e descontinuados
app.get('/api/saidas', async (req, res) => {
  const dados = await Saida.find({
    visivel: { $ne: false },
    descontinuado: { $ne: true }
  });
  res.json(dados);
});

// Marcar como descontinuado
app.post('/descontinuar/:id', async (req, res) => {
  try {
    const item = await Saida.findById(req.params.id);
    if (!item) return res.sendStatus(404);
    item.descontinuado = true;
    await item.save();
    io.emit('update');
    res.sendStatus(200);
  } catch {
    res.status(500).send('Erro ao processar descontinuação');
  }
});

// Adicionar item
app.post('/adicionar', async (req, res) => {
  const { nome, descricao, quantidade, unidade, tipo } = req.body;
  const data = getCurrentDate();
  const novoItem = new Saida({
    nome, descricao, quantidade, unidade, tipo,
    data, delivered: false, visivel: true, timestamp: Date.now()
  });
  await novoItem.save();
  io.emit('update');
  res.sendStatus(201);
});

// Inserir múltiplos itens
app.post('/itens', async (req, res) => {
  const { solicitante, destino, autorizado, count, ...rest } = req.body;
  const data = getCurrentDate();
  const timestamp = Date.now();
  const registros = [];
  for (let i = 0; i < Number(count); i++) {
    registros.push({
      nome: solicitante,
      observacao: destino,
      responsavel: autorizado,
      tipo: rest[`tipo_${i}`],
      descricao: rest[`descricao_${i}`],
      quantidade: rest[`quantidade_${i}`],
      unidade: rest[`unidade_${i}`],
      data,
      delivered: false,
      visivel: true,
      timestamp,
      descontinuado: false
    });
  }
  await Saida.insertMany(registros);
  io.emit('update');
  res.redirect('/solicitacao.html');
});

// Editar item
app.post('/editar/:id', async (req, res) => {
  try {
    const item = await Saida.findById(req.params.id);
    if (!item) return res.status(404).send('Item não encontrado');
    const { nome, descricao, quantidade, unidade, tipo, observacao } = req.body;
    Object.assign(item, { nome, descricao, quantidade, unidade, tipo });
    item.observacao = `${observacao} (editado em ${new Date().toLocaleString()})`;
    await item.save();
    io.emit('update');
    res.sendStatus(200);
  } catch {
    res.status(500).send('Erro ao editar item');
  }
});

// Exportar Excel (intervalo)
app.get('/exportar-relatorio-excel', async (req, res) => {
  const { inicio, fim } = req.query;
  if (!inicio || !fim) return res.status(400).send('Datas de início e fim são obrigatórias.');
  const dados = await Saida.find({ data: { $gte: inicio, $lte: fim } });
  if (!dados.length) return res.status(404).send('Nenhum item encontrado no intervalo informado.');

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Relatório');
  worksheet.columns = [
    { header: 'Nome', key: 'nome', width: 20 },
    { header: 'Descrição', key: 'descricao', width: 25 },
    { header: 'Quantidade', key: 'quantidade', width: 10 },
    { header: 'Unidade', key: 'unidade', width: 10 },
    { header: 'Tipo', key: 'tipo', width: 15 },
    { header: 'Data', key: 'data', width: 15 },
    { header: 'Entregue?', key: 'delivered', width: 10 },
    { header: 'Observação', key: 'observacao', width: 30 },
    { header: 'Descontinuado?', key: 'descontinuado', width: 15 }
  ];
  dados.forEach(i => worksheet.addRow({
    nome: i.nome,
    descricao: i.descricao,
    quantidade: i.quantidade,
    unidade: i.unidade,
    tipo: i.tipo,
    data: i.data,
    delivered: i.delivered ? 'Sim' : 'Não',
    observacao: i.observacao || '',
    descontinuado: i.descontinuado ? 'Sim' : 'Não'
  }));

  const fileName = `relatorio_${inicio}_a_${fim}.xlsx`;
  const filePath = path.join(__dirname, fileName);
  await workbook.xlsx.writeFile(filePath);
  res.download(filePath, fileName, err => { if (!err) fs.unlinkSync(filePath); });
});

// Exportar Relatório do Dia
app.get('/exportar-relatorio', async (req, res) => {
  const currentDate = getCurrentDate();
  const dados = await Saida.find({ data: currentDate });
  if (!dados.length) return res.status(404).send('Nenhum item encontrado para o dia atual');

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Relatório do Dia');
  worksheet.columns = [
    { header: 'Nome', key: 'nome', width: 20 },
    { header: 'Descrição', key: 'descricao', width: 25 },
    { header: 'Quantidade', key: 'quantidade', width: 10 },
    { header: 'Unidade', key: 'unidade', width: 10 },
    { header: 'Tipo', key: 'tipo', width: 15 },
    { header: 'Data', key: 'data', width: 15 },
    { header: 'Entregue?', key: 'delivered', width: 10 },
    { header: 'Observação', key: 'observacao', width: 30 },
    { header: 'Descontinuado?', key: 'descontinuado', width: 15 }
  ];
  dados.forEach(i => worksheet.addRow({
    nome: i.nome,
    descricao: i.descricao,
    quantidade: i.quantidade,
    unidade: i.unidade,
    tipo: i.tipo,
    data: i.data,
    delivered: i.delivered ? 'Sim' : 'Não',
    observacao: i.observacao || '',
    descontinuado: i.descontinuado ? 'Sim' : 'Não'
  }));

  const fileName = `relatorio_${currentDate}.xlsx`;
  const filePath = path.join(__dirname, fileName);
  await workbook.xlsx.writeFile(filePath);
  res.download(filePath, fileName, err => { if (!err) fs.unlinkSync(filePath); });
});

// Toggle delivered
app.post('/reordenar/:id', async (req, res) => {
  try {
    const item = await Saida.findById(req.params.id);
    if (!item) return res.sendStatus(404);
    item.delivered = !item.delivered;
    await item.save();
    io.emit('update');
    res.sendStatus(200);
  } catch {
    res.status(500).send('Erro ao marcar devolução');
  }
});

// Remoção lógica
app.post('/remover/:id', async (req, res) => {
  try {
    const item = await Saida.findById(req.params.id);
    if (!item) return res.sendStatus(404);
    item.visivel = false;
    await item.save();
    io.emit('update');
    res.sendStatus(200);
  } catch {
    res.status(500).send('Erro ao processar remoção');
  }
});

// Página inicial
app.get('/', (req, res) => res.redirect('/solicitacao.html'));

// WebSocket
io.on('connection', socket => {
  console.log('🔌 Cliente conectado via Socket.io');
  socket.on('disconnect', () => console.log('🔌 Cliente desconectado'));
});

// Inicia servidor
server.listen(PORT, () => console.log(`🚀 Servidor rodando na porta ${PORT}`));
