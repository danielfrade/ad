# Definir codificação UTF-8 para entrada e saída
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Script de gerenciamento do Active Directory
Import-Module ActiveDirectory

# Configuração do domínio

$domain = "" #  Dominio do local
$baseDN = ""  # Domínio base para construção dos caminhos

# Verifica se o script já está rodando como administrador
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Este script requer privilégios de administrador. Solicitando elevação..." -ForegroundColor Yellow
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -WindowStyle Normal"
    Exit
}

# Função para aplicar customizações ao console
function Set-ConsoleCustomization {
    $Host.UI.RawUI.ForegroundColor = "White"
    $Host.UI.RawUI.BackgroundColor = "Black"
    Clear-Host
}

# Aplica as customizações ao iniciar
Set-ConsoleCustomization

function Show-Menu {
    Clear-Host
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host "          MENU DE GERENCIAMENTO AD            " -ForegroundColor Yellow
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host " Domínio: $domain" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------"
    Write-Host "1 - Criar usuário no AD" -ForegroundColor White
    Write-Host "2 - Inativar usuário no AD" -ForegroundColor White
    Write-Host "3 - Reativar usuário no AD" -ForegroundColor White
    Write-Host "4 - Deletar usuário" -ForegroundColor White
    Write-Host "5 - Resetar a senha" -ForegroundColor White
    Write-Host "6 - Desbloquear usuário" -ForegroundColor White
    Write-Host "7 - Associar computador no AD" -ForegroundColor White
    Write-Host "8 - Desassociar computador" -ForegroundColor White
    Write-Host "9 - Deletar computador" -ForegroundColor White
    Write-Host "10 - Alterar ramal" -ForegroundColor White
    Write-Host "11 - Sincronizar AD" -ForegroundColor White
    Write-Host "12 - Listar usuários" -ForegroundColor White
    Write-Host "13 - Listar computadores" -ForegroundColor White
    Write-Host "14 - Mover objeto para outra OU" -ForegroundColor White
    Write-Host "15 - Adicionar usuário a um grupo" -ForegroundColor White
    Write-Host "16 - Remover usuário de um grupo" -ForegroundColor White
    Write-Host "17 - Verificar membros de um grupo" -ForegroundColor White
    Write-Host "18 - Alterar atributos de um usuário" -ForegroundColor White
    Write-Host "19 - Exportar relatório de usuários" -ForegroundColor White
    Write-Host "20 - Exportar relatório de computadores" -ForegroundColor White
    Write-Host "21 - Exportar relatório de grupos" -ForegroundColor White
    Write-Host "0 - Sair do script" -ForegroundColor Red
    Write-Host "==============================================" -ForegroundColor Cyan
}

function Get-FullOUPath {
    param (
        [string]$baseOUChoice,
        [string]$subOUName
    )
    switch ($baseOUChoice) {
        "1" { # Departamentos
            return "OU=$subOUName,OU=Departamentos,$baseDN"
        }
        "2" { # Consultoria
            return "OU=$subOUName,OU=Consultoria,$baseDN"
        }
        "3" { # Filiais
            return "OU=Usuarios,OU=$subOUName,OU=Filiais,$baseDN"
        }
        "4" { # Usuários de Serviços de TI
            return "OU=$subOUName,OU=Usuarios de Serviços de TI,$baseDN"
        }
        default {
            Write-Host "Opção de OU base inválida." -ForegroundColor Red
            return $null
        }
    }
}

function Create-User {
    $firstName = Read-Host "Digite o primeiro nome do usuário: "
    $lastName = Read-Host "Digite o sobrenome do usuário: "
    $matricula = Read-Host "Digite a matrícula Vilma: "
    $centroCusto = Read-Host "Digite o centro de custo: "
    $cargo = Read-Host "Digite o cargo na Vilma: "
    $ramal = Read-Host "Digite o ramal: "
    $email = Read-Host "Digite o e-mail Vilma: "
    $username = Read-Host "Digite o nome de usuário (login): "
    $password = Read-Host "Digite a senha: " -AsSecureString
    $description = Read-Host "Digite a descrição (cargo do colaborador): "

    # Escolha da OU base
    Write-Host "Escolha a OU base para o usuário:" -ForegroundColor Cyan
    Write-Host "1 - Departamentos" -ForegroundColor White
    Write-Host "2 - Consultoria" -ForegroundColor White
    Write-Host "3 - Filiais" -ForegroundColor White
    Write-Host "4 - Usuários de Serviços de TI" -ForegroundColor White
    $baseOUChoice = Read-Host "Digite o número correspondente: "

    # Solicitar a sub-OU com base na escolha
    if ($baseOUChoice -eq "1") {
        $subOUName = Read-Host "Digite o setor (ex: T.I, RH, Financeiro): "
    } elseif ($baseOUChoice -eq "2") {
        $subOUName = Read-Host "Digite o nome da empresa (ex: EmpresaX, EmpresaY): "
    } elseif ($baseOUChoice -eq "3") {
        $subOUName = Read-Host "Digite o nome da filial (ex: Bahia, Betim, Cambé): "
    } elseif ($baseOUChoice -eq "4") {
        $subOUName = Read-Host "Digite o nome da sub-OU (ex: BI, Conceito, SAP, RM): "
    } else {
        Write-Host "Opção inválida. Usuário não será criado." -ForegroundColor Red
        return
    }

    # Obter o caminho completo da OU
    $ouPath = Get-FullOUPath -baseOUChoice $baseOUChoice -subOUName $subOUName

    if ($ouPath) {
        try {
            New-ADUser -Name "$firstName $lastName" `
                       -DisplayName "$firstName $lastName" `
                       -GivenName $firstName `
                       -Surname $lastName `
                       -Initials $matricula `
                       -Office $centroCusto `
                       -Title $cargo `
                       -Description $description `
                       -Company "DOCI" `
                       -OfficePhone $ramal `
                       -EmailAddress $email `
                       -SamAccountName $username `
                       -UserPrincipalName "$username@$domain" `
                       -AccountPassword $password `
                       -Enabled $true `
                       -Path $ouPath `
                       -ErrorAction Stop
            Write-Host "Usuário $username criado com sucesso em $ouPath." -ForegroundColor Green
        } catch {
            Write-Host "Erro ao criar usuário: $_" -ForegroundColor Red
        }
    }
}

function Disable-User {
    $username = Read-Host "Digite o nome do usuário para inativar: "
    try {
        Disable-ADAccount -Identity $username -ErrorAction Stop
        Write-Host "Usuário $username inativado com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao inativar usuário: $_" -ForegroundColor Red
    }
}

function Enable-User {
    $username = Read-Host "Digite o nome do usuário para reativar: "
    try {
        Enable-ADAccount -Identity $username -ErrorAction Stop
        Write-Host "Usuário $username reativado com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao reativar usuário: $_" -ForegroundColor Red
    }
}

function Remove-User {
    $username = Read-Host "Digite o nome do usuário para deletar: "
    try {
        Remove-ADUser -Identity $username -Confirm:$false -ErrorAction Stop
        Write-Host "Usuário $username deletado com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao deletar usuário: $_" -ForegroundColor Red
    }
}

function Reset-Password {
    $username = Read-Host "Digite o nome do usuário para resetar a senha: "
    $newPassword = Read-Host "Digite a nova senha: " -AsSecureString
    try {
        Set-ADAccountPassword -Identity $username -NewPassword $newPassword -Reset -ErrorAction Stop
        Write-Host "Senha do usuário $username resetada com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao resetar senha: $_" -ForegroundColor Red
    }
}

function Unlock-User {
    $username = Read-Host "Digite o nome do usuário para desbloquear: "
    try {
        Unlock-ADAccount -Identity $username -ErrorAction Stop
        Write-Host "Usuário $username desbloqueado com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao desbloquear usuário: $_" -ForegroundColor Red
    }
}

function Add-Computer {
    $computername = Read-Host "Digite o nome do computador: "
    Write-Host "Escolha a OU base para o computador:" -ForegroundColor Cyan
    Write-Host "1 - Departamentos" -ForegroundColor White
    Write-Host "2 - Consultoria" -ForegroundColor White
    Write-Host "3 - Filiais" -ForegroundColor White
    Write-Host "4 - Usuários de Serviços de TI" -ForegroundColor White
    $baseOUChoice = Read-Host "Digite o número correspondente: "

    if ($baseOUChoice -eq "1") {
        $subOUName = Read-Host "Digite o setor (ex: T.I, RH, Financeiro): "
    } elseif ($baseOUChoice -eq "2") {
        $subOUName = Read-Host "Digite o nome da empresa (ex: EmpresaX, EmpresaY): "
    } elseif ($baseOUChoice -eq "3") {
        $subOUName = Read-Host "Digite o nome da filial (ex: Bahia, Betim, Cambé): "
    } elseif ($baseOUChoice -eq "4") {
        $subOUName = Read-Host "Digite o nome da sub-OU (ex: BI, Conceito, SAP, RM): "
    } else {
        Write-Host "Opção inválida. Computador não será associado." -ForegroundColor Red
        return
    }

    $ouPath = Get-FullOUPath -baseOUChoice $baseOUChoice -subOUName $subOUName

    if ($ouPath) {
        try {
            New-ADComputer -Name $computername -Path $ouPath -ErrorAction Stop
            Write-Host "Computador $computername associado com sucesso em $ouPath." -ForegroundColor Green
        } catch {
            Write-Host "Erro ao associar computador: $_" -ForegroundColor Red
        }
    }
}

function Remove-Computer {
    $computername = Read-Host "Digite o nome do computador para desassociar: "
    try {
        Remove-ADComputer -Identity $computername -Confirm:$false -ErrorAction Stop
        Write-Host "Computador $computername desassociado com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao desassociar computador: $_" -ForegroundColor Red
    }
}

function Delete-Computer {
    $computername = Read-Host "Digite o nome do computador para deletar: "
    try {
        Remove-ADComputer -Identity $computername -Confirm:$false -ErrorAction Stop
        Write-Host "Computador $computername deletado com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao deletar computador: $_" -ForegroundColor Red
    }
}

function Change-Extension {
    $username = Read-Host "Digite o nome do usuário: "
    $extension = Read-Host "Digite o novo ramal: "
    try {
        Set-ADUser -Identity $username -OfficePhone $extension -ErrorAction Stop
        Write-Host "Ramal do usuário $username alterado para $extension." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao alterar ramal: $_" -ForegroundColor Red
    }
}

function Sync-AD {
    param (
        [string]$ComputerName = "SRVAD01"
    )
    $credential = Get-Credential -Message "Digite suas credenciais para sincronizar o AD (ex: dominio\usuario)"
    Write-Host "Sincronizando AD..." -ForegroundColor Yellow
    try {
        Invoke-Command -ComputerName $ComputerName -Credential $credential -ScriptBlock {
            Import-Module ADSync -ErrorAction Stop
            Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
        } -ErrorAction Stop
        Write-Host "Sincronização concluída." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao sincronizar AD: $_" -ForegroundColor Red
    }
}

function List-Users {
    Write-Host "Escolha a OU base para listar usuários:" -ForegroundColor Cyan
    Write-Host "1 - Departamentos" -ForegroundColor White
    Write-Host "2 - Consultoria" -ForegroundColor White
    Write-Host "3 - Filiais" -ForegroundColor White
    Write-Host "4 - Usuários de Serviços de TI" -ForegroundColor White
    $baseOUChoice = Read-Host "Digite o número correspondente: "

    if ($baseOUChoice -eq "1") {
        $subOUName = Read-Host "Digite o setor (ex: T.I, RH, Financeiro): "
    } elseif ($baseOUChoice -eq "2") {
        $subOUName = Read-Host "Digite o nome da empresa (ex: EmpresaX, EmpresaY): "
    } elseif ($baseOUChoice -eq "3") {
        $subOUName = Read-Host "Digite o nome da filial (ex: Bahia, Betim, Cambé): "
    } elseif ($baseOUChoice -eq "4") {
        $subOUName = Read-Host "Digite o nome da sub-OU (ex: BI, Conceito, SAP, RM): "
    } else {
        Write-Host "Opção inválida." -ForegroundColor Red
        return
    }

    $ouPath = Get-FullOUPath -baseOUChoice $baseOUChoice -subOUName $subOUName

    if ($ouPath) {
        try {
            Get-ADUser -Filter * -SearchBase $ouPath -ErrorAction Stop | Select-Object Name, SamAccountName, Enabled | Format-Table -AutoSize
        } catch {
            Write-Host "Erro ao listar usuários: $_" -ForegroundColor Red
        }
    }
}

function List-Computers {
    Write-Host "Escolha a OU base para listar computadores:" -ForegroundColor Cyan
    Write-Host "1 - Departamentos" -ForegroundColor White
    Write-Host "2 - Consultoria" -ForegroundColor White
    Write-Host "3 - Filiais" -ForegroundColor White
    Write-Host "4 - Usuários de Serviços de TI" -ForegroundColor White
    $baseOUChoice = Read-Host "Digite o número correspondente: "

    if ($baseOUChoice -eq "1") {
        $subOUName = Read-Host "Digite o setor (ex: T.I, RH, Financeiro): "
    } elseif ($baseOUChoice -eq "2") {
        $subOUName = Read-Host "Digite o nome da empresa (ex: EmpresaX, EmpresaY): "
    } elseif ($baseOUChoice -eq "3") {
        $subOUName = Read-Host "Digite o nome da filial (ex: Bahia, Betim, Cambé): "
    } elseif ($baseOUChoice -eq "4") {
        $subOUName = Read-Host "Digite o nome da sub-OU (ex: BI, Conceito, SAP, RM): "
    } else {
        Write-Host "Opção inválida." -ForegroundColor Red
        return
    }

    $ouPath = Get-FullOUPath -baseOUChoice $baseOUChoice -subOUName $subOUName

    if ($ouPath) {
        try {
            Get-ADComputer -Filter * -SearchBase $ouPath -ErrorAction Stop | Select-Object Name, Enabled | Format-Table -AutoSize
        } catch {
            Write-Host "Erro ao listar computadores: $_" -ForegroundColor Red
        }
    }
}

function Move-Object {
    $object = Read-Host "Digite o nome do usuário ou computador: "
    Write-Host "Escolha a OU base de destino:" -ForegroundColor Cyan
    Write-Host "1 - Departamentos" -ForegroundColor White
    Write-Host "2 - Consultoria" -ForegroundColor White
    Write-Host "3 - Filiais" -ForegroundColor White
    Write-Host "4 - Usuários de Serviços de TI" -ForegroundColor White
    $baseOUChoice = Read-Host "Digite o número correspondente: "

    if ($baseOUChoice -eq "1") {
        $subOUName = Read-Host "Digite o setor (ex: T.I, RH, Financeiro): "
    } elseif ($baseOUChoice -eq "2") {
        $subOUName = Read-Host "Digite o nome da empresa (ex: EmpresaX, EmpresaY): "
    } elseif ($baseOUChoice -eq "3") {
        $subOUName = Read-Host "Digite o nome da filial (ex: Bahia, Betim, Cambé): "
    } elseif ($baseOUChoice -eq "4") {
        $subOUName = Read-Host "Digite o nome da sub-OU (ex: BI, Conceito, SAP, RM): "
    } else {
        Write-Host "Opção inválida." -ForegroundColor Red
        return
    }

    $newOU = Get-FullOUPath -baseOUChoice $baseOUChoice -subOUName $subOUName

    if ($newOU) {
        try {
            Get-ADObject -Filter { Name -eq $object } -ErrorAction Stop | Move-ADObject -TargetPath $newOU -ErrorAction Stop
            Write-Host "Objeto $object movido para $newOU com sucesso." -ForegroundColor Green
        } catch {
            Write-Host "Erro ao mover objeto: $_" -ForegroundColor Red
        }
    }
}

function Add-UserToGroup {
    $username = Read-Host "Digite o nome do usuário: "
    $group = Read-Host "Digite o nome do grupo: "
    try {
        $user = Get-ADUser -Identity $username -ErrorAction Stop
        $groupObj = Get-ADGroup -Identity $group -ErrorAction Stop
        Add-ADGroupMember -Identity $group -Members $username -ErrorAction Stop
        Write-Host "Usuário $username adicionado ao grupo $group com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao adicionar usuário ao grupo: $_" -ForegroundColor Red
    }
}

function Remove-UserFromGroup {
    $username = Read-Host "Digite o nome do usuário: "
    $group = Read-Host "Digite o nome do grupo: "
    try {
        Remove-ADGroupMember -Identity $group -Members $username -Confirm:$false -ErrorAction Stop
        Write-Host "Usuário $username removido do grupo $group com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao remover usuário do grupo: $_" -ForegroundColor Red
    }
}

function Get-GroupMembers {
    $group = Read-Host "Digite o nome do grupo: "
    try {
        Get-ADGroupMember -Identity $group -ErrorAction Stop | Select-Object Name, SamAccountName | Format-Table -AutoSize
    } catch {
        Write-Host "Erro ao listar membros do grupo: $_" -ForegroundColor Red
    }
}

function Set-UserAttributes {
    $username = Read-Host "Digite o nome do usuário: "
    $attribute = Read-Host "Digite o atributo a ser alterado (ex: Title, Department): "
    $value = Read-Host "Digite o novo valor: "
    try {
        Set-ADUser -Identity $username -Replace @{ $attribute = $value } -ErrorAction Stop
        Write-Host "Atributo $attribute do usuário $username alterado para $value com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao alterar atributo: $_" -ForegroundColor Red
    }
}

function Export-UserReport {
    Write-Host "Escolha a OU base para exportar usuários:" -ForegroundColor Cyan
    Write-Host "1 - Departamentos" -ForegroundColor White
    Write-Host "2 - Consultoria" -ForegroundColor White
    Write-Host "3 - Filiais" -ForegroundColor White
    Write-Host "4 - Usuários de Serviços de TI" -ForegroundColor White
    $baseOUChoice = Read-Host "Digite o número correspondente: "

    if ($baseOUChoice -eq "1") {
        $subOUName = Read-Host "Digite o setor (ex: T.I, RH, Financeiro): "
    } elseif ($baseOUChoice -eq "2") {
        $subOUName = Read-Host "Digite o nome da empresa (ex: EmpresaX, EmpresaY): "
    } elseif ($baseOUChoice -eq "3") {
        $subOUName = Read-Host "Digite o nome da filial (ex: Bahia, Betim, Cambé): "
    } elseif ($baseOUChoice -eq "4") {
        $subOUName = Read-Host "Digite o nome da sub-OU (ex: BI, Conceito, SAP, RM): "
    } else {
        Write-Host "Opção inválida." -ForegroundColor Red
        return
    }

    $ouPath = Get-FullOUPath -baseOUChoice $baseOUChoice -subOUName $subOUName
    $outputFile = Read-Host "Digite o nome do arquivo de saída (ex: usuários.csv): "

    if ($ouPath) {
        try {
            Get-ADUser -Filter * -SearchBase $ouPath -ErrorAction Stop | Select-Object Name, SamAccountName, Enabled | Export-Csv -Path $outputFile -NoTypeInformation -ErrorAction Stop
            Write-Host "Relatório de usuários exportado para $outputFile com sucesso." -ForegroundColor Green
        } catch {
            Write-Host "Erro ao exportar relatório: $_" -ForegroundColor Red
        }
    }
}

function Export-ComputerReport {
    Write-Host "Escolha a OU base para exportar computadores:" -ForegroundColor Cyan
    Write-Host "1 - Departamentos" -ForegroundColor White
    Write-Host "2 - Consultoria" -ForegroundColor White
    Write-Host "3 - Filiais" -ForegroundColor White
    Write-Host "4 - Usuários de Serviços de TI" -ForegroundColor White
    $baseOUChoice = Read-Host "Digite o número correspondente: "

    if ($baseOUChoice -eq "1") {
        $subOUName = Read-Host "Digite o setor (ex: T.I, RH, Financeiro): "
    } elseif ($baseOUChoice -eq "2") {
        $subOUName = Read-Host "Digite o nome da empresa (ex: EmpresaX, EmpresaY): "
    } elseif ($baseOUChoice -eq "3") {
        $subOUName = Read-Host "Digite o nome da filial (ex: Bahia, Betim, Cambé): "
    } elseif ($baseOUChoice -eq "4") {
        $subOUName = Read-Host "Digite o nome da sub-OU (ex: BI, Conceito, SAP, RM): "
    } else {
        Write-Host "Opção inválida." -ForegroundColor Red
        return
    }

    $ouPath = Get-FullOUPath -baseOUChoice $baseOUChoice -subOUName $subOUName
    $outputFile = Read-Host "Digite o nome do arquivo de saída (ex: computadores.csv): "

    if ($ouPath) {
        try {
            Get-ADComputer -Filter * -SearchBase $ouPath -ErrorAction Stop | Select-Object Name, Enabled, LastLogonDate | Export-Csv -Path $outputFile -NoTypeInformation -ErrorAction Stop
            Write-Host "Relatório de computadores exportado para $outputFile com sucesso." -ForegroundColor Green
        } catch {
            Write-Host "Erro ao exportar relatório: $_" -ForegroundColor Red
        }
    }
}

function Export-GroupReport {
    $setor = Read-Host "Digite o setor (ex: T.I, RH, Financeiro): "
    $ouPath = "OU=Grupos,OU=$setor,OU=Departamentos,$baseDN"
    $outputFile = Read-Host "Digite o nome do arquivo de saída (ex: grupos.csv): "
    try {
        Get-ADGroup -Filter * -SearchBase $ouPath -ErrorAction Stop | 
        Select-Object Name, SamAccountName, GroupCategory, GroupScope | 
        Export-Csv -Path $outputFile -NoTypeInformation -ErrorAction Stop
        Write-Host "Relatório de grupos exportado para $outputFile com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao exportar relatório de grupos: $_" -ForegroundColor Red
    }
}

# Loop do menu
do {
    Show-Menu
    $input = Read-Host "Digite o número correspondente à opção desejada: "
    switch ($input) {
        '1' { Create-User }
        '2' { Disable-User }
        '3' { Enable-User }
        '4' { Remove-User }
        '5' { Reset-Password }
        '6' { Unlock-User }
        '7' { Add-Computer }
        '8' { Remove-Computer }
        '9' { Delete-Computer }
        '10' { Change-Extension }
        '11' { Sync-AD }
        '12' { List-Users }
        '13' { List-Computers }
        '14' { Move-Object }
        '15' { Add-UserToGroup }
        '16' { Remove-UserFromGroup }
        '17' { Get-GroupMembers }
        '18' { Set-UserAttributes }
        '19' { Export-UserReport }
        '20' { Export-ComputerReport }
        '21' { Export-GroupReport }
        '0' { Write-Host "Saindo do script..." -ForegroundColor Red }
        default { Write-Host "Opção inválida, tente novamente." -ForegroundColor Red }
    }
    if ($input -ne '0') {
        Write-Host "Pressione Enter para continuar..." -ForegroundColor Gray
        $null = Read-Host
    }
} until ($input -eq '0')
