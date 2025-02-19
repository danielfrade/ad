# Gerenciamento de Active Directory com PowerShell

Criar e gerenciar usuários de forma ágil, personalizando atributos e permissões em segundos.
Resetar senhas e desbloquear contas com apenas um clique, sem a necessidade de navegar por menus complexos.
Associar e desassociar computadores ao AD de forma rápida e organizada, mantendo seu inventário sempre atualizado.
Sincronizar o AD para garantir que todas as alterações sejam aplicadas em tempo real, evitando inconsistências.
Benefícios que vão além da praticidade
Automatize tarefas repetitivas: Libere-se de trabalhos manuais e repetitivos, ganhando tempo para atividades estratégicas.
Reduza erros humanos: A automação garante a consistência e precisão das tarefas, diminuindo a chance de falhas.
Aumente a eficiência: Execute tarefas complexas em segundos, otimizando o tempo da sua equipe.
Centralize o gerenciamento: Controle múltiplos dispositivos e usuários de forma centralizada, simplificando a administração da rede.
Prepare-se para o futuro: A escalabilidade do script permite que ele se adapte às necessidades de redes maiores e mais complexas.

---

## Funcionalidades

O script oferece as seguintes funcionalidades:

- **Criação de Usuários**: Cria usuários em massa a partir de um arquivo CSV ou TXT.
- **Inativação/Reativação de Contas**: Desabilita ou reativa contas de usuários.
- **Reset de Senhas**: Reseta a senha de um usuário.
- **Desbloqueio de Contas**: Desbloqueia contas de usuários.
- **Associação de Computadores**: Adiciona computadores ao domínio.
- **Movimentação de Objetos**: Move usuários ou computadores entre OUs.
- **Exportação de Relatórios**: Gera relatórios de usuários e computadores em formato CSV.
- **Gerenciamento de Grupos**: Adiciona ou remove usuários de grupos.
- **Sincronização do AD**: Executa a sincronização do Active Directory.

---

## Como Usar

### Pré-requisitos

- **PowerShell**: O script requer o PowerShell 5.1 ou superior.
- **Módulo ActiveDirectory**: Certifique-se de que o módulo do Active Directory está instalado.
  ```powershell
  Install-WindowsFeature -Name RSAT-AD-PowerShell
  ```

### Executando o Script

1. **Clone o Repositório**:
   ```bash
   git clone https://github.com/seu-usuario/seu-repositorio.git
   ```

2. **Execute o Script**:
   - Abra o PowerShell como administrador.
   - Navegue até o diretório do script.
   - Execute o script:
     ```powershell
     .\Vilma.ps1
     ```

3. **Siga o Menu**:
   - O script exibirá um menu interativo com todas as opções disponíveis.
   - Escolha a opção desejada e siga as instruções.

---

## Estrutura do Projeto

- **GerenciamentoAD.ps1**: Script principal com o menu e todas as funcionalidades.
- **README.md**: Este arquivo, com a documentação do projeto.

---

## Exemplos de Uso

### Criando Usuários

1. Execute o script e escolha a opção Criar usuário
2. Insira o nome do usuário que deseja criar com os dados respectivos.

### Inativando um Usuário

1. Execute o script e escolha a opção **2 - Inativar usuário no AD**.
2. Insira o nome do usuário que deseja inativar.

---

## Contribuindo

Contribuições são bem-vindas! Siga os passos abaixo:

1. Faça um fork do projeto.
2. Crie uma branch para sua feature (`git checkout -b feature/nova-feature`).
3. Commit suas mudanças (`git commit -m 'Adicionando nova feature'`).
4. Faça push para a branch (`git push origin feature/nova-feature`).
5. Abra um Pull Request.
