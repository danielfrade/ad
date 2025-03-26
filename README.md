# Gerenciamento de Active Directory com PowerShell

Automatize e simplifique a administração do Active Directory. Crie usuários, gerencie permissões, resete senhas e sincronize alterações em tempo real com eficiência e precisão.

### Por que usar?
- **Automatização**: Elimine tarefas repetitivas e ganhe tempo.  
- **Precisão**: Reduza erros com processos consistentes.  
- **Eficiência**: Execute ações complexas em segundos.  
- **Centralização**: Controle usuários e dispositivos em um só lugar.  
- **Escalabilidade**: Adapte-se a redes de qualquer tamanho.

---

## Funcionalidades

- Criação de usuários em massa (CSV/TXT)  
- Ativação/inativação de contas  
- Reset e desbloqueio de senhas  
- Associação de computadores ao domínio  
- Movimentação de objetos entre OUs  
- Relatórios em CSV  
- Gerenciamento de grupos  
- Sincronização do AD  

---

## Como Usar

### Pré-requisitos
- PowerShell 5.1+  
- Módulo ActiveDirectory:  
  ```powershell
  Install-WindowsFeature -Name RSAT-AD-PowerShell
  ```

### Instalação
1. Clone o repositório:  
   ```bash
   git clone https://github.com/danielfrade/ad
   ```
2. Execute como administrador:  
   ```powershell
   .\ActiveDirectory.ps1
   ```
3. Use o menu interativo para navegar.

---

## Estrutura
- `ActiveDirectory.ps1`: Script principal  
- `README.md`: Documentação  

---

## Exemplos
- **Criar usuário**: Selecione a opção e insira os dados.  
- **Inativar conta**: Escolha a opção e informe o usuário.  

---

## Contribuições
Quer ajudar? Faça um fork, crie uma branch (`git checkout -b feature/nova-ideia`), commit suas mudanças e envie um Pull Request!
