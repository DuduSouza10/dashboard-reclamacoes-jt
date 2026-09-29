# Controle de Reclamação Atendimento - Railway

Dashboard web baseado na aba `Dashboard` do arquivo original e alimentado pela base `STATUS PEDIDO`.

## O que já vem pronto
- Dashboard público somente leitura.
- Filtros por período, Regional, Cliente, Base, Supervisor, RM, Dupla, Retorno, Anexo e Status.
- KPIs: Pedidos, Valor da mercadoria, Entregue, Pendente de atualização, Extravio e Taxa de reversão.
- Gráficos de status, evolução diária e bases com maior volume.
- Tabelas por RM e por Dupla.
- Área `Assistentes` protegida por senha para editar a base.
- Edição de Status, Anexo e Retorno; criação e exclusão de casos.
- Importação de XLSX/CSV em modo mesclar ou substituir.
- Exportação CSV dos registros filtrados.
- Base inicial já carregada a partir do arquivo fornecido.
- Português / 中文 no dashboard.

## Deploy recomendado no Railway
1. Crie um novo projeto a partir deste diretório/repositório.
2. Adicione um serviço **PostgreSQL** ao projeto.
3. No serviço do dashboard, confirme que `DATABASE_URL` foi disponibilizada pelo Railway.
4. Configure as variáveis:
   - `SECRET_KEY`: uma chave longa e aleatória.
   - `ASSISTANT_PASSWORD`: senha que só os assistentes conhecerão.
5. Faça o deploy. O comando de inicialização já está no `railway.json`/`Procfile`.

Na primeira inicialização, se a tabela estiver vazia, o sistema importa `seed/initial_data.csv` automaticamente.

## Teste local
```bash
python -m venv .venv
.venv\\Scripts\\activate
pip install -r requirements.txt
set PORT=5000
python app.py
```
Acesse `http://localhost:5000`.

## Observação sobre persistência
No Railway, use PostgreSQL. O fallback SQLite serve para teste local; sem Volume, o filesystem do serviço pode ser recriado em novos deploys.
