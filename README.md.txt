# SICOT
### Sistema de Cruzamento Operacional Telefônico e Telemático

---

## 📌 Descrição

O **SICOT** é uma ferramenta desenvolvida para auxiliar na análise operacional de dados provenientes de:

- Interceptações telefônicas (fixo e móvel)
- Interceptações telemáticas (WhatsApp)
- Cruzamento com dados cadastrais de operadoras

O sistema permite:

- 📊 Geração de estatísticas
- 🔎 Cruzamento de dados
- 🏆 Ranking de terminais por volume de contatos
- 📤 Exportação de relatórios (Excel / TXT)
- 📥 Importação de dados cadastrais

---

## ⚙️ Tecnologias Utilizadas

- Python 3.12+
- Streamlit
- Pandas
- OpenPyXL

---

## 🚀 Como Executar Localmente

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run SICOT_v1.0.3_PRODUCAO.py