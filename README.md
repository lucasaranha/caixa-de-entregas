# Caixa de Entregas

Página para os alunos entregarem atividades com nome, série e arquivo (PDF, Word,
PowerPoint ou imagem). O aluno recebe um comprovante com protocolo na hora.

- **A página** fica no GitHub Pages (`index.html` + `config.js`).
- **Os arquivos** vão para o Google Drive do professor pelo Google Apps Script
  (`apps-script/Code.gs`): uma pasta **Caixa de Entregas** com uma subpasta por série
  e a planilha **Caixa de Entregas · registro**, com uma linha por entrega e uma
  caixinha "Corrigida".

```
index.html            página dos alunos
config.js             link do servidor (Apps Script)
apps-script/Code.gs   servidor: salva no Drive e anota na planilha
apps-script/appsscript.json
```

## Ligar o servidor (uma vez)

1. Abra https://script.google.com › **Novo projeto** e renomeie para **Caixa de Entregas**.
2. No arquivo `Código.gs`, apague tudo e cole o conteúdo de `apps-script/Code.gs`. Salve.
3. No topo, escolha a função **configurar** › **Executar** › **Revisar permissões**.
   Quando aparecer "O Google não verificou este app": **Avançado** ›
   **Acessar Caixa de Entregas (não seguro)** › **Permitir**. É normal, porque o app é seu.
4. **Implantar** › **Nova implantação** › engrenagem › **App da Web**.
   Executar como: **Eu**. Quem pode acessar: **Qualquer pessoa**. **Implantar**.
5. Copie o **URL do app da Web** (termina em `/exec`) e cole em `config.js`, entre as aspas de `servidor`.

Para conferir: abrir o link `/exec` no navegador mostra `{"ok":true,...}`.

## Mudanças depois

- **Mudou o Code.gs?** Implantar › Gerenciar implantações › lápis › Versão: **Nova versão** › Implantar.
  O link `/exec` continua o mesmo.
- **Trocar o bimestre:** em `index.html`, procure `4º bimestre · 2026`.
- **Séries:** a lista fica em `index.html` e em `SERIES` no `Code.gs`. As chaves precisam ser iguais.
- **Limites:** 10 MB por arquivo e 5 arquivos por entrega. Ficam em `LIMITE_MB` e `MAX_ARQUIVOS`
  no `Code.gs` e no `index.html`, e o texto "Até 10 MB" aparece na página.
- Qualquer pessoa com o link consegue enviar. Divulgue só para as turmas.
