Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$FR_Main                         = New-Object system.Windows.Forms.Form
$FR_Main.ClientSize              = New-Object System.Drawing.Point(515,577)
$FR_Main.text                    = "Copiator Tajabara"
$FR_Main.TopMost                 = $false

$LB_Titulo                       = New-Object system.Windows.Forms.Label
$LB_Titulo.text                  = "Copiator"
$LB_Titulo.AutoSize              = $true
$LB_Titulo.width                 = 25
$LB_Titulo.height                = 10
$LB_Titulo.Anchor                = 'top'
$LB_Titulo.location              = New-Object System.Drawing.Point(231,15)
$LB_Titulo.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$GB_Dispositivos                 = New-Object system.Windows.Forms.Groupbox
$GB_Dispositivos.height          = 116
$GB_Dispositivos.width           = 487
$GB_Dispositivos.Anchor          = 'top'
$GB_Dispositivos.text            = "Dispositivos"
$GB_Dispositivos.location        = New-Object System.Drawing.Point(16,36)

$LB_Dispositivo                  = New-Object system.Windows.Forms.Label
$LB_Dispositivo.text             = "Dispositivos"
$LB_Dispositivo.AutoSize         = $true
$LB_Dispositivo.width            = 25
$LB_Dispositivo.height           = 10
$LB_Dispositivo.Anchor           = 'top'
$LB_Dispositivo.location         = New-Object System.Drawing.Point(14,31)
$LB_Dispositivo.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$CB_Dispositivos                 = New-Object system.Windows.Forms.ComboBox
$CB_Dispositivos.width           = 265
$CB_Dispositivos.height          = 20
$CB_Dispositivos.enabled         = $true
$CB_Dispositivos.Anchor          = 'top'
$CB_Dispositivos.location        = New-Object System.Drawing.Point(112,31)
$CB_Dispositivos.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$BT_Refresh                      = New-Object system.Windows.Forms.Button
$BT_Refresh.text                 = "Refresh"
$BT_Refresh.width                = 66
$BT_Refresh.height               = 30
$BT_Refresh.Anchor               = 'top'
$BT_Refresh.location             = New-Object System.Drawing.Point(405,33)
$BT_Refresh.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$GB_Dados                        = New-Object system.Windows.Forms.Groupbox
$GB_Dados.height                 = 191
$GB_Dados.width                  = 487
$GB_Dados.Anchor                 = 'top'
$GB_Dados.text                   = "Informações"
$GB_Dados.location               = New-Object System.Drawing.Point(15,288)

$LB_Nome                         = New-Object system.Windows.Forms.Label
$LB_Nome.text                    = "Nome"
$LB_Nome.AutoSize                = $true
$LB_Nome.width                   = 25
$LB_Nome.height                  = 10
$LB_Nome.location                = New-Object System.Drawing.Point(13,37)
$LB_Nome.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TB_Nome                         = New-Object system.Windows.Forms.TextBox
$TB_Nome.multiline               = $false
$TB_Nome.width                   = 380
$TB_Nome.height                  = 20
$TB_Nome.location                = New-Object System.Drawing.Point(82,32)
$TB_Nome.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$LB_Data                         = New-Object system.Windows.Forms.Label
$LB_Data.text                    = "Data"
$LB_Data.AutoSize                = $true
$LB_Data.width                   = 25
$LB_Data.height                  = 10
$LB_Data.location                = New-Object System.Drawing.Point(13,78)
$LB_Data.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TM_Data                         = New-Object system.Windows.Forms.MaskedTextBox
$TM_Data.multiline               = $false
$TM_Data.width                   = 102
$TM_Data.height                  = 20
$TM_Data.mask                    = "00/00/0000"
$TM_Data.location                = New-Object System.Drawing.Point(82,78)
$TM_Data.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$LB_Inicio                       = New-Object system.Windows.Forms.Label
$LB_Inicio.text                  = "Inicio"
$LB_Inicio.AutoSize              = $true
$LB_Inicio.width                 = 25
$LB_Inicio.height                = 10
$LB_Inicio.location              = New-Object System.Drawing.Point(13,122)
$LB_Inicio.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TM_Inicio                       = New-Object system.Windows.Forms.MaskedTextBox
$TM_Inicio.multiline             = $false
$TM_Inicio.width                 = 100
$TM_Inicio.height                = 20
$TM_Inicio.mask                  = "00:00:00"
$TM_Inicio.location              = New-Object System.Drawing.Point(82,122)
$TM_Inicio.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$LB_Fim                          = New-Object system.Windows.Forms.Label
$LB_Fim.text                     = "Fim"
$LB_Fim.AutoSize                 = $true
$LB_Fim.width                    = 25
$LB_Fim.height                   = 10
$LB_Fim.location                 = New-Object System.Drawing.Point(13,166)
$LB_Fim.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TM_Fim                          = New-Object system.Windows.Forms.MaskedTextBox
$TM_Fim.multiline                = $false
$TM_Fim.width                    = 100
$TM_Fim.height                   = 20
$TM_Fim.mask                     = "00:00:00"
$TM_Fim.location                 = New-Object System.Drawing.Point(82,166)
$TM_Fim.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$BT_Executa                      = New-Object system.Windows.Forms.Button
$BT_Executa.text                 = "Executar"
$BT_Executa.width                = 94
$BT_Executa.height               = 30
$BT_Executa.Anchor               = 'top'
$BT_Executa.location             = New-Object System.Drawing.Point(202,518)
$BT_Executa.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$GB_Localizacao                 = New-Object system.Windows.Forms.Groupbox
$GB_Localizacao.height          = 79
$GB_Localizacao.width           = 485
$GB_Localizacao.text            = "Localização"
$GB_Localizacao.Anchor          = 'top'
$GB_Localizacao.location        = New-Object System.Drawing.Point(15,178)

$LB_Caminho                      = New-Object system.Windows.Forms.Label
$LB_Caminho.text                 = "Caminho"
$LB_Caminho.AutoSize             = $true
$LB_Caminho.width                = 25
$LB_Caminho.height               = 10
$LB_Caminho.Anchor               = 'top'
$LB_Caminho.location             = New-Object System.Drawing.Point(14,33)
$LB_Caminho.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TP_Localizacao                  = New-Object system.Windows.Forms.ToolTip
$TP_Localizacao.ToolTipTitle     = "Localização"
$TP_Localizacao.isBalloon        = $true

$TB_Caminho                      = New-Object system.Windows.Forms.TextBox
$TB_Caminho.multiline            = $false
$TB_Caminho.width                = 267
$TB_Caminho.height               = 20
$TB_Caminho.enabled              = $true
$TB_Caminho.Anchor               = 'top'
$TB_Caminho.location             = New-Object System.Drawing.Point(110,32)
$TB_Caminho.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$BT_SelecionPasta                = New-Object system.Windows.Forms.Button
$BT_SelecionPasta.text           = "..."
$BT_SelecionPasta.width          = 60
$BT_SelecionPasta.height         = 30
$BT_SelecionPasta.Anchor         = 'top'
$BT_SelecionPasta.location       = New-Object System.Drawing.Point(406,33)
$BT_SelecionPasta.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$LB_Memoria                      = New-Object system.Windows.Forms.Label
$LB_Memoria.text                 = "Memoria"
$LB_Memoria.AutoSize             = $true
$LB_Memoria.width                = 25
$LB_Memoria.height               = 10
$LB_Memoria.location             = New-Object System.Drawing.Point(16,75)
$LB_Memoria.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$CB_Memoria                      = New-Object system.Windows.Forms.ComboBox
$CB_Memoria.width                = 265
$CB_Memoria.height               = 20
$CB_Memoria.location             = New-Object System.Drawing.Point(112,75)
$CB_Memoria.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TP_Localizacao.SetToolTip($LB_Caminho,'Local onde serão salvas as fotos')
$TP_Localizacao.SetToolTip($TB_Caminho,'Local onde serão salvas as fotos')
$TP_Localizacao.SetToolTip($BT_SelecionPasta,'Local onde serão salvas as fotos')
$FR_Main.controls.AddRange(@($LB_Titulo,$GB_Dispositivos,$GB_Dados,$BT_Executa,$GB_Localizacao))
$GB_Dispositivos.controls.AddRange(@($LB_Dispositivo,$CB_Dispositivos,$BT_Refresh,$LB_Memoria,$CB_Memoria))
$GB_Dados.controls.AddRange(@($LB_Nome,$TB_Nome,$LB_Data,$TM_Data,$LB_Inicio,$TM_Inicio,$LB_Fim,$TM_Fim))
$GB_Localizacao.controls.AddRange(@($LB_Caminho,$TB_Caminho,$BT_SelecionPasta))

$BT_Refresh.Add_Click({ Set-TodosDiretorios })
$BT_SelecionPasta.Add_Click({ Get-FolderDestination })
$BT_Executa.Add_Click({ Set-Start })
$CB_Dispositivos.Add_TextChanged({ Set-Armazenamentos })
$FR_Main.Add_Load({ Set-TodosDiretorios })


function Create-Dir($path)
{
  if(! (Test-Path -Path $path))
  {
    New-Item -Path $path -ItemType Directory
  }
}


function Get-AbreCaixadePasta()
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Escolha a pasta"
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK")
    {
        return $foldername.SelectedPath
    }
    
}

function Get-FolderDestination($initialDirectory="")
{
    $Dispositivo = $CB_Dispositivos.Text
    $Memoria = $CB_Memoria.Text
    $TB_Caminho.Text = Get-AbreCaixadePasta
    $CB_Dispositivos.Text = $Dispositivo
    $CB_Memoria.Text = $Memoria
}


function Get-SubFolder($parentDir, $subPath)
{
  $result = $parentDir
  foreach($pathSegment in ($subPath -split "\\"))
  {
    $result = $result.GetFolder.Items() | Where-Object {$_.Name -eq $pathSegment} | select -First 1
    if($result -eq $null)
    {
      throw "Not found $subPath folder"
    }
  }
  return $result
}


function Get-PhoneMainDir($phoneName)
{
    $o = New-Object -com Shell.Application
    $rootComputerDirectory = $o.NameSpace(0x11)
    if ($phoneName -eq 'all') {
       $phoneDirectory = $rootComputerDirectory.Items()
    } else {
       $phoneDirectory = $rootComputerDirectory.Items() | Where-Object {$_.Name -eq $phoneName} | select -First 1
    }

    return $phoneDirectory
}


function Set-TodosDiretorios()
{
    $todosDispositivos = Get-PhoneMainDir "all"
    $CB_Dispositivos.Items.Clear()
    $CB_Dispositivos.Text = ""
    $CB_Memoria.Items.Clear()
    $CB_Memoria.Text = ""
    $quantidade = 0
    foreach($dispositivo in $todosDispositivos) {
        if ($dispositivo.Type -eq "Media Player Portátil") {
            $CB_Dispositivos.Items.Add($dispositivo.Name)
            $quantidade += 1
        }
    }

    if ($quantidade -eq 0) {
        $CB_Dispositivos.Items.Add("Nenhum dispositivo conectado")
    }
 
}

function Set-Armazenamentos()
{

    if (($CB_Dispositivos.Text -ne "") -and ($CB_Dispositivos.Text -ne "Nenhum dispositivo conectado")) {
        $CB_Memoria.Items.Clear()
        $CB_Memoria.Text = ""
        $quantidade = 0

        $sourceMtpDir = Get-PhoneMainDir $CB_Dispositivos.Text
        $fullSourceDirPath = Get-FullPathOfMtpDir $sourceMtpDir

        foreach ($item in $sourceMtpDir.GetFolder.Items()) {
            if($item.IsFolder) {
                $CB_Memoria.Items.Add($item.name)
                $quantidade += 1
        
            }
        }
        if ($quantidade -eq 0) {
            $CB_Memoria.Items.Add("Selecione um dispositivo")
        }
    }
}


function Get-Arquivo($itemName, $expression = '')
{
    if ($expression -eq '') {
        $filter='(.jpg)|(.mp4)$'
    } else {
        $filter = $expression
    }
    return $item.Name -match $filter
}


function Get-DataHoraNomeArquivo([String]$itemName) 
{
   try{
        $t = $itemName -match '(\d{8}_\d{6})'

        $dataRAW = $Matches[1].Substring(0,8)
        $horaRAW = $Matches[1].Substring(9,6)

        $ano = $dataRAW.Substring(0,4)
        $mes = $dataRAW.Substring(4,2)
        $dia = $dataRAW.Substring(6,2)
        $horas = $horaRAW.Substring(0,2)
        $minutos = $horaRAW.Substring(2,2)
        $segundos = $horaRAW.Substring(4,2)

        $dataCompleta = -join($ano, "-", $mes, "-", $dia, " ", $horas, ":", $minutos, ":", $segundos)

        return Get-Date $dataCompleta
    } catch {
        return $false
    } 
}


function Get-DataUsuario([String]$dataUsuario)
{

   try{
        $Data = @((Get-Date $dataUsuario).ToString("yyyy"), (Get-Date $dataUsuario).toString("MM"), (Get-Date $dataUsuario).toString("dd"))
        return $Data
    } catch {
        return $false
    }
    
}


function Get-FullPathOfMtpDir($mtpDir)
{
     $fullDirPath = ""
     $directory = $mtpDir.GetFolder
     while($directory -ne $null)
     {
       $fullDirPath =  -join($directory.Title, '\', $fullDirPath)
       $directory = $directory.ParentFolder
     }
     return $fullDirPath
}


function Copy-FromPhoneSource-ToBackup($sourceMtpDir, $destDirPath, $HoraInicio, $HoraFim)
{

    Create-Dir $destDirPath
    $destDirShell = (new-object -com Shell.Application).NameSpace($destDirPath)
    $fullSourceDirPath = Get-FullPathOfMtpDir $sourceMtpDir

    foreach ($item in $sourceMtpDir.GetFolder.Items())
    {
        #Write-Host ($item | Format-List | Out-String)

        if($item.IsFolder) {

            Copy-FromPhoneSource-ToBackup $item $destDirPath $HoraInicio $HoraFim

        } else {

            if ((Get-Arquivo $item.Name) -eq $true) {
               
                $PossuiData = Get-DataHoraNomeArquivo $item.Name
                if ($PossuiData -ne $false) {
                            
                    if (($PossuiData -ge $HoraInicio) -and ($PossuiData -le $HoraFim)) {
                        Write-Host $item.Name
                        $destDirShell.CopyHere($item)
                    }
                }
            }

        }

    }

}

function Set-Start ()
{

    $NomeTelefone = Get-PhoneMainDir $CB_Dispositivos.Text
    $LocalTelefone = $CB_Memoria.Text

    $DataUsuario = Get-DataUsuario($TM_Data.Text)

    $CaminhoLocal = -join ($TB_Caminho.Text, "\", $TB_Nome.Text, "\", $DataUsuario[0], "\", $DataUsuario[1], "\", $DataUsuario[2])
    
    $HoraInicio = Get-Date (-join ($DataUsuario[0], "-", $DataUsuario[1], "-", $DataUsuario[2], " ", $TM_Inicio.Text))
    $HoraFim = Get-Date (-join ($DataUsuario[0], "-", $DataUsuario[1], "-", $DataUsuario[2], " ", $TM_Fim.Text))

    Copy-FromPhoneSource-ToBackup (Get-SubFolder $NomeTelefone $LocalTelefone) $CaminhoLocal $HoraInicio $HoraFim
}


[void]$FR_Main.ShowDialog()