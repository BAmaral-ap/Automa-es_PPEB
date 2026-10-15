import importlib.util
from pathlib import Path
import subprocess
import tempfile
import unittest


spec = importlib.util.spec_from_file_location(
    'atualizador', Path(__file__).resolve().parents[1] / 'atualizar_repositorio.py')
atualizador = importlib.util.module_from_spec(spec)
spec.loader.exec_module(atualizador)


class AtualizacaoTest(unittest.TestCase):
    def comando(self, *args):
        return subprocess.run(['git', *map(str, args)], check=True, capture_output=True,
                              text=True).stdout.strip()

    def git(self, repo, *args):
        return self.comando('-C', repo, *args)

    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        raiz = Path(temp.name)
        self.remoto = raiz / 'remoto.git'
        self.autor = raiz / 'autor'
        self.local = raiz / 'local'
        self.comando('init', '--bare', '-b', 'main', self.remoto)
        self.comando('clone', self.remoto, self.autor)
        self.identidade(self.autor)
        (self.autor / 'arquivo.txt').write_text('inicial', encoding='utf-8')
        self.git(self.autor, 'add', '.')
        self.git(self.autor, 'commit', '-m', 'Inicial')
        self.git(self.autor, 'push', 'origin', 'main')
        self.comando('clone', self.remoto, self.local)
        self.identidade(self.local)

    def identidade(self, repo):
        self.git(repo, 'config', 'user.name', 'Teste')
        self.git(repo, 'config', 'user.email', 'teste@example.invalid')

    def publicar(self):
        (self.autor / 'arquivo.txt').write_text('atualizado', encoding='utf-8')
        self.git(self.autor, 'commit', '-am', 'Atualização remota')
        self.git(self.autor, 'push', 'origin', 'main')

    def test_aplica_novo_commit_e_pode_repetir(self):
        self.publicar()
        self.assertIn('Atualizado para o commit', atualizador.atualizar(self.local))
        self.assertEqual((self.local / 'arquivo.txt').read_text(encoding='utf-8'), 'atualizado')
        self.assertEqual(self.git(self.local, 'rev-parse', 'HEAD'),
                         self.git(self.autor, 'rev-parse', 'HEAD'))
        self.assertIn('nenhuma novidade', atualizador.atualizar(self.local))

    def test_preserva_edicoes_locais(self):
        self.publicar()
        antes = self.git(self.local, 'rev-parse', 'HEAD')
        (self.local / 'arquivo.txt').write_text('meu trabalho', encoding='utf-8')
        self.assertIn('alterações locais', atualizador.atualizar(self.local))
        self.assertEqual(self.git(self.local, 'rev-parse', 'HEAD'), antes)
        self.assertEqual((self.local / 'arquivo.txt').read_text(encoding='utf-8'), 'meu trabalho')

    def test_preserva_arquivos_novos(self):
        self.publicar()
        (self.local / 'novo.txt').write_text('meu arquivo', encoding='utf-8')
        self.assertIn('alterações locais', atualizador.atualizar(self.local))
        self.assertEqual((self.local / 'arquivo.txt').read_text(encoding='utf-8'), 'inicial')
        self.assertEqual((self.local / 'novo.txt').read_text(encoding='utf-8'), 'meu arquivo')

    def test_nao_troca_ou_atualiza_outra_branch(self):
        self.publicar()
        self.git(self.local, 'switch', '-c', 'minha-automacao')
        self.assertIn('selecione a branch main', atualizador.atualizar(self.local))
        self.assertEqual(self.git(self.local, 'branch', '--show-current'), 'minha-automacao')
        self.assertEqual((self.local / 'arquivo.txt').read_text(encoding='utf-8'), 'inicial')

    def test_preserva_commits_divergentes_sem_criar_merge(self):
        self.publicar()
        (self.local / 'meu.txt').write_text('meu commit', encoding='utf-8')
        self.git(self.local, 'add', '.')
        self.git(self.local, 'commit', '-m', 'Trabalho local')
        antes = self.git(self.local, 'rev-parse', 'HEAD')
        self.assertIn('commits locais', atualizador.atualizar(self.local))
        self.assertEqual(self.git(self.local, 'rev-parse', 'HEAD'), antes)
        self.assertEqual(self.git(self.local, 'status', '--porcelain'), '')

    def test_nao_interfere_em_operacao_do_git(self):
        self.publicar()
        (self.local / '.git' / 'rebase-merge').mkdir()
        self.assertIn('operação do Git', atualizador.atualizar(self.local))
        self.assertEqual((self.local / 'arquivo.txt').read_text(encoding='utf-8'), 'inicial')


if __name__ == '__main__':
    unittest.main()
