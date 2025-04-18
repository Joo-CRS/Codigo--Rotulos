📌 Gerador Automático de Rótulos em PDF

Script Python que automatiza a criação de rótulos em PDF a partir de um modelo Word, inserindo automaticamente:

Data de fabricação (A partir de 01/01/2025)
Data de validade (EX: 60 dias após fabricação)
Código de lote no formato LYYDDD076YY

🛠 Funcionamento:
✔ Gera 1 PDF por dia (365 dias)
✔ Nomeia arquivos no padrão: (Ex: Rótulo_Pó_AAAA-MM-DD.pdf)
✔ Código de lote automático (ex: L25001076YY para 01/01/2025)

⚠️ Solução de problemas:
Se travar, feche o Word manualmente
Verifique se o caminho do modelo está correto
Execute como administrador se necessário

📂 Saída:
/Teste_arquivos
├── Rótulo_Pó_2025-01-01.docx
├── Rótulo_Pó_2025-01-01.pdf
├── Rótulo_Pó_2025-01-02.docx

📝 Observações:
Compatível com Windows (requer MS Word)
Pode ser adaptado para outros tipos de documentos
(Pó é pó de malte haha)
