const express = require('express');
const mongoose = require('mongoose');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

const mongoUrl = process.env.MONGO_URL;
mongoose.connect(mongoUrl)
  .then(() => {
    console.log('Conectado ao MongoDB Atlas!');
    // Inicie o servidor só após conectar ao banco
    app.listen(3001, () => {
      console.log('Servidor rodando na porta 3001');
    });
  })
  .catch((err) => {
    console.error('Erro ao conectar ao MongoDB:', err);
  });

const PesquisaSchema = new mongoose.Schema({}, { strict: false });
const Pesquisa = mongoose.model('Pesquisa', PesquisaSchema);

app.post('/pesquisas', async (req, res) => {
  const pesquisa = new Pesquisa(req.body);
  await pesquisa.save();
  res.status(201).json(pesquisa);
});

app.get('/pesquisas', async (req, res) => {
  const pesquisas = await Pesquisa.find();
  res.json(pesquisas);
});