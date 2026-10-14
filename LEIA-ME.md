# Finanças pessoais (Apps Script + Google Sheets)

## Instalar
1. Crie uma planilha nova no Google Sheets (ex.: "Finanças").
2. Extensões > Apps Script. Crie os arquivos Logic, Data, Setup e Api (tipo Script) e Index (tipo HTML),
   colando o conteúdo de cada arquivo desta pasta. Em Configurações do projeto, marque "Mostrar appsscript.json"
   e cole o conteúdo do appsscript.json.
3. Recarregue a planilha. No menu Finanças > Configurar planilha (aceite as permissões).
4. Na aba Contas: preencha dia de fechamento, vencimento e conta de pagamento dos cartões e, se quiser, o saldo inicial.
5. Implantar > Nova implantação > App da Web. Executar como: eu. Quem tem acesso: somente eu.
6. Abra o endereço no celular e no computador, entre com sua conta Google e salve nos favoritos / tela inicial.

## Importar o histórico
Cole as abas antigas (com o cabeçalho) em Importar, ImportarSaldos e ImportarRecorrentes e use o menu Finanças.
Pode repetir: o que já existe não duplica. Confira a aba DePara (CC, CV e CM sem número viram qual conta?).

## Atualização para a versão 4 (grupos, orçamento, painel, a receber)

1. Cole os arquivos Logic.gs, Data.gs, Setup.gs, Api.gs e Index.html no Apps Script.
2. Salve, volte à planilha, recarregue a página e use o menu Finanças > "Atualizar para a versão 4". Pode rodar mais de uma vez.
3. Implantar > Gerenciar implantações > lápis > Nova versão > Implantar.

## Backup

Menu Finanças: "Fazer backup agora", "Ligar backup semanal" e "Desligar backup semanal". As cópias ficam na pasta "Backups Finanças", ao lado da planilha, e o app guarda as últimas 12. Na primeira vez o Google pede autorização para acessar o Drive. A aba Mais do app avisa se o backup está desligado ou falhou.
