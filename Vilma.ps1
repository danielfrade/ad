# Certifique-se de salvar este arquivo com codificação UTF-8 com BOM
Import-Module ActiveDirectory

# Configuração do domínio
#Digite o dominio respectivo
$domain = ""
#Digite sua OU respectiva BASE
$baseOU = ""  # Caminho base da OU

function Show-Menu {
    Clear-Host
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host "          MENU DE GERENCIAMENTO AD            " -ForegroundColor Yellow
    Write-Host "==============================================" -ForegroundColor Cyan
    Write-Host " Domínio: $domain" -ForegroundColor Green
    Write-Host "--------------------------------------------------------------"
    Write-Host "1 - Criar usuario no AD" -ForegroundColor White
    Write-Host "2 - Inativar usuario no AD" -ForegroundColor White
    Write-Host "3 - Reativar usuario no AD" -ForegroundColor White
    Write-Host "4 - Deletar usuario" -ForegroundColor White
    Write-Host "5 - Resetar a senha" -ForegroundColor White
    Write-Host "6 - Desbloquear usuario" -ForegroundColor White
    Write-Host "7 - Associar computador no AD" -ForegroundColor White
    Write-Host "8 - Desassociar computador" -ForegroundColor White
    Write-Host "9 - Deletar computador" -ForegroundColor White
    Write-Host "10 - Alterar ramal" -ForegroundColor White
    Write-Host "11 - Sincronizar AD" -ForegroundColor White
    Write-Host "12 - Listar usuarios" -ForegroundColor White
    Write-Host "13 - Listar computadores" -ForegroundColor White
    Write-Host "14 - Mover objeto para outra OU" -ForegroundColor White
    Write-Host "15 - Adicionar usuario a um grupo" -ForegroundColor White
    Write-Host "16 - Remover usuario de um grupo" -ForegroundColor White
    Write-Host "17 - Verificar membros de um grupo" -ForegroundColor White
    Write-Host "18 - Alterar atributos de um usuario" -ForegroundColor White
    Write-Host "19 - Exportar relatorio de usuarios" -ForegroundColor White
    Write-Host "20 - Exportar relatorio de computadores" -ForegroundColor White
	Write-Host "21 - Exportar relatorio de grupos" -ForegroundColor White
    Write-Host "0 - Sair do script" -ForegroundColor Red
    Write-Host "==============================================" -ForegroundColor Cyan
}

function Get-FullOUPath {
    param (
        [string]$ouName
    )
    return "OU=$ouName,$baseOU"
}

function Create-User {
    $firstName = Read-Host "Digite o primeiro nome do usuario: "
    $lastName = Read-Host "Digite o sobrenome do usuario: "
    $matricula = Read-Host "Digite a matricula Vilma: "
    $centroCusto = Read-Host "Digite o centro de custo: "
    $cargo = Read-Host "Digite o cargo na Vilma: "
    $ramal = Read-Host "Digite o ramal: "
    $email = Read-Host "Digite o e-mail Vilma: "
    $username = Read-Host "Digite o nome de usuario (login): "
    $password = Read-Host "Digite a senha: " -AsSecureString
    $setor = Read-Host "Digite o setor (ex: T.I, RH, Financeiro): "
    $ouPath = "OU=Usuarios,OU=$setor,$baseOU"  # Caminho da OU Usuarios dentro do setor

    try {
        New-ADUser -Name "$firstName $lastName" `
                   -DisplayName "$firstName $lastName" `
                   -GivenName $firstName `
                   -Surname $lastName `
                   -Initials $matricula `
                   -Office $centroCusto `
                   -Description $cargo `
                   -OfficePhone $ramal `
                   -EmailAddress $email `
                   -SamAccountName $username `
                   -AccountPassword $password `
                   -Enabled $true `
                   -Path $ouPath `
                   -ErrorAction Stop
        Write-Host "Usuario $username criado com sucesso em $ouPath." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao criar usuario: $_" -ForegroundColor Red
    }
}

function Disable-User {
    $username = Read-Host "Digite o nome do usuario para inativar: "
    try {
        Disable-ADAccount -Identity $username -ErrorAction Stop
        Write-Host "Usuario $username inativado com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao inativar usuario: $_" -ForegroundColor Red
    }
}

function Enable-User {
    $username = Read-Host "Digite o nome do usuario para reativar: "
    try {
        Enable-ADAccount -Identity $username -ErrorAction Stop
        Write-Host "Usuario $username reativado com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao reativar usuario: $_" -ForegroundColor Red
    }
}

function Remove-User {
    $username = Read-Host "Digite o nome do usuario para deletar: "
    try {
        Remove-ADUser -Identity $username -Confirm:$false -ErrorAction Stop
        Write-Host "Usuario $username deletado com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao deletar usuario: $_" -ForegroundColor Red
    }
}

function Reset-Password {
    $username = Read-Host "Digite o nome do usuario para resetar a senha: "
    $newPassword = Read-Host "Digite a nova senha: " -AsSecureString
    try {
        Set-ADAccountPassword -Identity $username -NewPassword $newPassword -Reset -ErrorAction Stop
        Write-Host "Senha do usuario $username resetada com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao resetar senha: $_" -ForegroundColor Red
    }
}

function Unlock-User {
    $username = Read-Host "Digite o nome do usuario para desbloquear: "
    try {
        Unlock-ADAccount -Identity $username -ErrorAction Stop
        Write-Host "Usuario $username desbloqueado com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao desbloquear usuario: $_" -ForegroundColor Red
    }
}

function Add-Computer {
    $computername = Read-Host "Digite o nome do computador: "
    $ouName = Read-Host "Digite o nome da OU (ex: Massas): "
    $ouPath = Get-FullOUPath -ouName $ouName
    try {
        New-ADComputer -Name $computername -Path $ouPath -ErrorAction Stop
        Write-Host "Computador $computername associado com sucesso em $ouPath." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao associar computador: $_" -ForegroundColor Red
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
    $username = Read-Host "Digite o nome do usuario: "
    $extension = Read-Host "Digite o novo ramal: "
    try {
        Set-ADUser -Identity $username -OfficePhone $extension -ErrorAction Stop
        Write-Host "Ramal do usuario $username alterado para $extension." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao alterar ramal: $_" -ForegroundColor Red
    }
}

function Sync-AD {
    Write-Host "Sincronizando AD..." -ForegroundColor Yellow
    try {
        Invoke-Command -ScriptBlock { ipconfig /registerdns } -ErrorAction Stop
        Write-Host "Sincronização concluida." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao sincronizar AD: $_" -ForegroundColor Red
    }
}

function List-Users {
    $ouName = Read-Host "Digite o nome da OU (ex: Massas): "
    $ouPath = Get-FullOUPath -ouName $ouName
    try {
        Get-ADUser -Filter * -SearchBase $ouPath -ErrorAction Stop | Select-Object Name, SamAccountName, Enabled | Format-Table -AutoSize
    } catch {
        Write-Host "Erro ao listar usuarios: $_" -ForegroundColor Red
    }
}

function List-Computers {
    $ouName = Read-Host "Digite o nome da OU (ex: Massas): "
    $ouPath = Get-FullOUPath -ouName $ouName
    try {
        Get-ADComputer -Filter * -SearchBase $ouPath -ErrorAction Stop | Select-Object Name, Enabled | Format-Table -AutoSize
    } catch {
        Write-Host "Erro ao listar computadores: $_" -ForegroundColor Red
    }
}

function Move-Object {
    $object = Read-Host "Digite o nome do usuario ou computador: "
    $ouName = Read-Host "Digite o nome da nova OU (ex: Massas): "
    $newOU = Get-FullOUPath -ouName $ouName
    try {
        Get-ADObject -Filter { Name -eq $object } -ErrorAction Stop | Move-ADObject -TargetPath $newOU -ErrorAction Stop
        Write-Host "Objeto $object movido para $newOU com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao mover objeto: $_" -ForegroundColor Red
    }
}

function Add-UserToGroup {
    $username = Read-Host "Digite o nome do usuario: "
    $group = Read-Host "Digite o nome do grupo: "
    try {
        # Verifica se o usuário existe
        $user = Get-ADUser -Identity $username -ErrorAction Stop
        # Verifica se o grupo existe
        $groupObj = Get-ADGroup -Identity $group -ErrorAction Stop
        # Adiciona o usuário ao grupo
        Add-ADGroupMember -Identity $group -Members $username -ErrorAction Stop
        Write-Host "Usuario $username adicionado ao grupo $group com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao adicionar usuario ao grupo: $_" -ForegroundColor Red
    }
}

function Remove-UserFromGroup {
    $username = Read-Host "Digite o nome do usuario: "
    $group = Read-Host "Digite o nome do grupo: "
    try {
        Remove-ADGroupMember -Identity $group -Members $username -Confirm:$false -ErrorAction Stop
        Write-Host "Usuario $username removido do grupo $group com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao remover usuario do grupo: $_" -ForegroundColor Red
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
    $username = Read-Host "Digite o nome do usuario: "
    $attribute = Read-Host "Digite o atributo a ser alterado (ex: Title, Department): "
    $value = Read-Host "Digite o novo valor: "
    try {
        Set-ADUser -Identity $username -Replace @{ $attribute = $value } -ErrorAction Stop
        Write-Host "Atributo $attribute do usuario $username alterado para $value com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao alterar atributo: $_" -ForegroundColor Red
    }
}

function Export-UserReport {
    $ouName = Read-Host "Digite o nome da OU (ex: Massas): "
    $ouPath = Get-FullOUPath -ouName $ouName
    $outputFile = Read-Host "Digite o nome do arquivo de saída (ex: usuarios.csv): "
    try {
        Get-ADUser -Filter * -SearchBase $ouPath -ErrorAction Stop | Select-Object Name, SamAccountName, Enabled | Export-Csv -Path $outputFile -NoTypeInformation -ErrorAction Stop
        Write-Host "Relatório de usuarios exportado para $outputFile com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao exportar relatório: $_" -ForegroundColor Red
    }
}

function Export-ComputerReport {
    $ouName = Read-Host "Digite o nome da OU (ex: Massas): "
    $ouPath = Get-FullOUPath -ouName $ouName
    $outputFile = Read-Host "Digite o nome do arquivo de saída (ex: computadores.csv): "
    try {
        Get-ADComputer -Filter * -SearchBase $ouPath -ErrorAction Stop | Select-Object Name, Enabled, LastLogonDate | Export-Csv -Path $outputFile -NoTypeInformation -ErrorAction Stop
        Write-Host "Relatório de computadores exportado para $outputFile com sucesso." -ForegroundColor Green
    } catch {
        Write-Host "Erro ao exportar relatório: $_" -ForegroundColor Red
    }
}

function Export-GroupReport {
    $setor = Read-Host "Digite o setor (ex: T.I, RH, Financeiro): "
    $ouPath = "OU=Grupos,OU=$setor,$baseOU"  # Caminho da OU Grupos dentro do setor
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
    $input = Read-Host "Digite o numero correspondente a opcao desejada: "
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
