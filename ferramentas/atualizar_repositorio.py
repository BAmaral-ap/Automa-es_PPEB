"""Atualiza main a cada minuto, sem descartar arquivos ou commits locais."""
import argparse
import os
from pathlib import Path
import subprocess
import time


def git(repo, *args, verificar=True):
    ambiente = dict(os.environ, GIT_TERMINAL_PROMPT="0", GCM_INTERACTIVE="never",
                    GIT_MERGE_AUTOEDIT="no")
    return subprocess.run(["git", "-C", str(repo), *args], env=ambiente,
                          capture_output=True, text=True, encoding="utf-8", errors="replace",
                          timeout=45, check=verificar)


def impedimento(repo):
    if git(repo, "branch", "--show-current").stdout.strip() != "main":
        return "Pausado: selecione a branch main para receber atualizações."
    if git(repo, "status", "--porcelain", "--untracked-files=normal").stdout.strip():
        return "Pausado: há alterações locais. Faça commit ou guarde seu trabalho antes de atualizar."
    for nome in ("MERGE_HEAD", "CHERRY_PICK_HEAD", "REVERT_HEAD", "rebase-merge",
                 "rebase-apply", "sequencer", "BISECT_LOG", "index.lock", "HEAD.lock"):
        caminho = Path(git(repo, "rev-parse", "--git-path", nome).stdout.strip())
        if not caminho.is_absolute():
            caminho = repo / caminho
        if caminho.exists():
            return "Pausado: outra operação do Git está em andamento."
    return None


def atualizar(repo):
    try:
        motivo = impedimento(repo)
        if motivo:
            return motivo
        git(repo, "fetch", "origin", "refs/heads/main")
        # Reconfere depois da rede: o usuário pode ter começado a trabalhar.
        motivo = impedimento(repo)
        if motivo:
            return motivo
        atual = git(repo, "rev-parse", "HEAD").stdout.strip()
        novo = git(repo, "rev-parse", "FETCH_HEAD").stdout.strip()
        if atual == novo:
            return "Atualizado: nenhuma novidade no GitHub."
        if git(repo, "merge-base", "--is-ancestor", atual, novo, verificar=False).returncode != 0:
            return "Pausado: main tem commits locais. Sincronize manualmente pelo VS Code."
        git(repo, "merge", "--ff-only", novo)
        return f"Atualizado para o commit {novo[:7]}."
    except FileNotFoundError:
        return "Falha: Git não encontrado. Instale o Git e reabra o VS Code."
    except subprocess.TimeoutExpired:
        return "Falha: tempo de espera excedido. Tentarei novamente no próximo minuto."
    except subprocess.CalledProcessError:
        return "Falha no Git. Confira a conexão e execute git pull --ff-only origin main para ver o erro."


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--uma-vez", action="store_true", help="Verifica uma vez e encerra")
    args = parser.parse_args()
    repo = Path(__file__).resolve().parents[1]
    ultimo = None
    try:
        while True:
            mensagem = atualizar(repo)
            if mensagem != ultimo:
                print("[GitHub] " + mensagem, flush=True)
                ultimo = mensagem
            if args.uma_vez:
                return
            time.sleep(60)
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
