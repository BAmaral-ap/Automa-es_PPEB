# Atualização automática no VS Code

Com a pasta raiz `Automa-es_PPEB` aberta no VS Code, a tarefa
**Atualizar do GitHub automaticamente** verifica a branch `main` do GitHub a cada
minuto e aplica novos commits com avanço direto (`fast-forward`). Usa Python 3
e Git já instalados; no Windows, usa o comando `py -3`, como o gerador de cartazes.

## Ativação no Windows

1. Na cópia local, selecione `main` e receba esta configuração pelo terminal:

   ```powershell
   git switch main
   git pull --ff-only origin main
   ```

2. Abra no VS Code a pasta raiz do repositório, onde fica `.vscode`.
3. Pressione `Ctrl+Shift+P`, procure **Tasks: Manage Automatic Tasks in Folder**
   (Gerenciar Tarefas Automáticas na Pasta) e escolha **Allow Automatic Tasks**.
   Faça isso somente para este repositório de confiança.
4. Reabra a pasta ou execute **Developer: Reload Window** pela paleta.

A tarefa permanece ativa enquanto a pasta estiver aberta no VS Code. Na primeira
vez, faça um pull manual para concluir eventual login do GitHub; a tarefa não
abre diálogos de autenticação. As mensagens aparecem no terminal da tarefa.

Ela pausa quando existem arquivos alterados ou novos, commits locais, outra
branch selecionada ou uma operação do Git em andamento. Não faz push, reset,
stash nem troca de branch. Quando os impedimentos desaparecem, verifica de novo
automaticamente. Arquivos ignorados pelo Git não impedem a atualização.
Alterações ainda não salvas no editor não aparecem no Git: salve seu trabalho.

Para interromper, use **Tasks: Terminate Task**. Para desativar nas próximas
aberturas, use **Tasks: Manage Automatic Tasks in Folder → Disallow Automatic Tasks**.
Para conferir apenas uma vez:

```powershell
py -3 ferramentas/atualizar_repositorio.py --uma-vez
```

## Validação do atualizador

Os testes usam repositórios Git temporários, sem acessar o GitHub:

```powershell
py -3 -m unittest discover -s ferramentas/tests -v
```
