  $MouseEventSig = @'
        [DllImport("user32.dll",CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
        public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
'@

#Situarse en la carpeta del script, importante debido a los .py
Set-Location "C:\Users\PHINX\Desktop\TFG"

#Cargar los espacios de nombres System.Speech que contienen los tipos que admiten el reconocimiento de voz
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Speech")
 
#Crear un objeto de tipo motor de síntesis de voz que permite a PowerShell escuchar voz para una gramática con restricciones
$speechRecogEngRestr = [System.Speech.Recognition.SpeechRecognitionEngine]::new()
 
#Crear las restricciones para una gramática de reconocimiento de voz
#Añadir palabras de un diccionario
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("gato"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("notepad"))
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("cerrar notepad"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("champ"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("crome"))
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("cerrar crome"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("desfragmentador"))
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("ce"))
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("de"))
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("analizar"))
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("optimizar"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("apaga"))
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("reinicia"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("cisco"))
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("cerrar cisco"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("calendario"))
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("cerrar calendario"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("maximiza"))
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("minimiza"))
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("cierra"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("disco duro"))  
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #OK
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("explorador"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("maquina virtual"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("nueva"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("panel de control"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("cortafuegos"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("registro"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #OK
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("adaptador"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #OK
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("arbol"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #OK
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("andel"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("parar"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("arriba"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("abajo"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("izquierda"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("derecha"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("escape"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("aceptar"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.LoadGrammar([System.Speech.Recognition.GrammarBuilder]::new("tabular"))
#- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - #
$speechRecogEngRestr.SetInputToDefaultAudioDevice()




$Builder = New-BTContentBuilder
$Builder | Add-BTText "Bienvenido!","Siempre estoy camuflado, pero te escucho!"
Add-BTHeroImage -ContentBuilder $Builder -Source 'C:\Users\PHINX\Desktop\TFG\iconos\cat2.gif'
$Builder | Show-BTNotification

##------------------------------------------------------------------------------------------------
 do
{
    $orden = $speechRecogEngRestr.Recognize()
    $orden.text

    #NOTEPAD######################################################################################

     if($orden.text -eq "notepad") ##INICIA NOTEPAD
    {
     $Builder = New-BTContentBuilder
     $Builder | Add-BTText "Abriendo notepad"
     $Builder | Add-BTAppLogo "C:\Users\PHINX\Desktop\TFG\iconos\notepad.ico"
     $Builder | Show-BTNotification
     Start-Sleep 1
     Start-Process notepad}


    #XAMPP########################################################################################

        if($orden.text -eq "champ") ##INICIA XAMPP
    {
     $Builder = New-BTContentBuilder
     $Builder | Add-BTText "Iniciando Xampp"
     $Builder | Add-BTAppLogo "C:\Users\PHINX\Desktop\TFG\iconos\xampp.png"
     $Builder | Show-BTNotification
        Start-Sleep 1
            Start-Process "D:\xampp\xampp-control.exe"}


        if($orden.text -eq "sql") ##INICIA SQL
    {
     $Builder = New-BTContentBuilder
     $Builder | Add-BTText "Iniciando SQL"
     $Builder | Add-BTAppLogo "C:\Users\PHINX\Desktop\TFG\iconos\xampp.png"
     $Builder | Show-BTNotification
        Start-Sleep 1
            [System.Windows.Forms.SendKeys]::SendWait("{TAB}");start-sleep 1; 
            [System.Windows.Forms.SendKeys]::SendWait("{TAB}");start-sleep 1;
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}");start-sleep 1; 
            [System.Windows.Forms.SendKeys]::SendWait("{TAB}");start-sleep 1;
            [System.Windows.Forms.SendKeys]::SendWait("{TAB}");start-sleep 1;
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}");start-sleep 1;
            [System.Windows.Forms.SendKeys]::SendWait("{TAB}");start-sleep 1; 
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}");Start-Sleep 7; 
            [System.Windows.Forms.SendKeys]::SendWait("{TAB}");start-sleep 1; 
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")}


    #CHROME#######################################################################################
    
        if($orden.text -eq "crome") ##INICIA CHROME
    {
     $Builder = New-BTContentBuilder
     $Builder | Add-BTText "Iniciando"
     $Builder | Add-BTImage "C:\Users\PHINX\Desktop\TFG\iconos\chrome2.png"
     $Builder | Show-BTNotification
        Start-Sleep 1
            Start-Process "C:\Program Files\Google\Chrome\Application\chrome.exe"}


        if($orden.text -eq "cerrar crome") ##CIERRA CHROME
    {
     $Builder = New-BTContentBuilder
     $Builder | Add-BTText "Cerrando"
     $Builder | Add-BTImage "C:\Users\PHINX\Desktop\TFG\iconos\chrome2.png"
     $Builder | Show-BTNotification 
         Start-Sleep 1
             Stop-Process -name chrome}


    #DESFRAGMENTADOR##############################################################################

        if($orden.text -eq "desfragmentador") ##INICIA EL DESFRAGMENTADOR
    {
     $Builder = New-BTContentBuilder
     $Builder | Add-BTText "Iniciando desfragmentador"
     $Builder | Add-BTAppLogo "C:\Users\PHINX\Desktop\TFG\iconos\desfragmentador.png"
     $Builder | Show-BTNotification
        Start-Sleep 1
            Start-Process "C:\Windows\system32\dfrgui.exe"}


        if($orden.text -eq "ce") ##SELECCIONA LA UNIDAD C:
    {
     $Builder = New-BTContentBuilder
     $Builder | Add-BTText "Unidad C:"
     $Builder | Add-BTAppLogo "C:\Users\PHINX\Desktop\TFG\iconos\unidad.png"
     $Builder | Show-BTNotification
         Start-Sleep 1
             $posicion = python c.py
                $posicion = $posicion.replace("Box(","").replace(")","").split(",").trim()
                $MouseEvent = Add-Type -memberDefinition $MouseEventSig -name "MouseEventWinApi" -passThru
                    Start-Sleep -Seconds 1
                    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(([Int]$posicion[0].replace("left=","")+15),([Int]$posicion[1].replace("top=","")+15))
                    $MouseEvent::mouse_event(0x00000002, 0, 0, 0, 0)
                    $MouseEvent::mouse_event(0x00000004, 0, 0, 0, 0)}


        if($orden.text -eq "de") ##SELECCIONA LA UNIDAD D:
    {
     $Builder = New-BTContentBuilder
     $Builder | Add-BTText "Unidad D:"
     $Builder | Add-BTAppLogo "C:\Users\PHINX\Desktop\TFG\iconos\unidad.png"
     $Builder | Show-BTNotification
        Start-Sleep 1
             $posicion = python d.py
                $posicion = $posicion.replace("Box(","").replace(")","").split(",").trim()
                $MouseEvent = Add-Type -memberDefinition $MouseEventSig -name "MouseEventWinApi" -passThru
                    Start-Sleep -Seconds 1
                    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(([Int]$posicion[0].replace("left=","")+15),([Int]$posicion[1].replace("top=","")+15))
                    $MouseEvent::mouse_event(0x00000002, 0, 0, 0, 0)
                    $MouseEvent::mouse_event(0x00000004, 0, 0, 0, 0)}


        if($orden.text -eq "analizar")  ##ANALIZA LA UNIDAD
    {
     Start-Sleep 1
     $posicion = python desfrag-analiza.py
        $posicion = $posicion.replace("Box(","").replace(")","").split(",").trim()
        $MouseEvent = Add-Type -memberDefinition $MouseEventSig -name "MouseEventWinApi" -passThru
            Start-Sleep -Seconds 1
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(([Int]$posicion[0].replace("left=","")+15),([Int]$posicion[1].replace("top=","")+15))
            $MouseEvent::mouse_event(0x00000002, 0, 0, 0, 0)
            $MouseEvent::mouse_event(0x00000004, 0, 0, 0, 0)}


        if($orden.text -eq "optimizar") ##OPTIMIZA LA UNIDAD SELECCIONADA
    {
         Start-Sleep 1
            $posicion = python desfrag-optimiza.py
            $posicion = $posicion.replace("Box(","").replace(")","").split(",").trim()
            $MouseEvent = Add-Type -memberDefinition $MouseEventSig -name "MouseEventWinApi" -passThru
                Start-Sleep -Seconds 1
                [ System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(([Int]$posicion[0].replace("left=","")+15),([Int]$posicion[1].replace("top=","")+15))
                $MouseEvent::mouse_event(0x00000002, 0, 0, 0, 0)
                $MouseEvent::mouse_event(0x00000004, 0, 0, 0, 0)}


    #APAGAR/REINICIAR#############################################################################

        #if($orden.text -eq "apaga") ##APAGA EL PC
   # {
   #  $test = New-BTProgressBar -Status ‘_______________________________________________________’ -Indeterminate
   #  New-BurntToastNotification -ProgressBar $test -Text ‘Apagado en curso’ -AppLogo 'C:\Users\PHINX\Desktop\TFG\iconos\apagado-reinicio.ico' 
   #     Start-Sleep 5
   #         Stop-Computer -ComputerName localhost}


        #if($orden.text -eq "reinicia") ##REINICIA EL PC
    #{
   #  $test = New-BTProgressBar -Status ‘_______________________________________________________’ -Indeterminate
    # New-BurntToastNotification -ProgressBar $test -Text ‘Reinicio en curso’ -AppLogo 'C:\Users\PHINX\Desktop\TFG\iconos\apagado-reinicio.ico' 
    #    Start-Sleep 5 
     #   Start-Sleep 5
     #       Restart-Computer -ComputerName localhost}
    

    #CISCO########################################################################################

           if($orden.text -eq "cisco") ##INICIA CISCO
    {
     $Builder = New-BTContentBuilder
     $Builder | Add-BTText "Iniciando Cisco","Espere un segundo"
     $Builder | Add-BTImage "C:\Users\PHINX\Desktop\TFG\iconos\cisco.png"
     $Builder | Show-BTNotification
        Start-Sleep 1
            Start-Process "C:\Program Files (x86)\Cisco Packet Tracer 6.1.1sv\bin\PacketTracer6.exe"}


        if($orden.text -eq "cerrar cisco") ##CIERRA CISCO
    {
     $Builder = New-BTContentBuilder
     $Builder | Add-BTText "Cerrando Cisco"
     $Builder | Add-BTImage "C:\Users\PHINX\Desktop\TFG\iconos\cisco.png"
     $Builder | Show-BTNotification
        Start-Sleep 1
            Stop-Process -name PacketTracer6}


    #CALENDARIO###################################################################################

        if($orden.text -eq "calendario") ##ABRE EL CALENDARIO
    {
     $Builder = New-BTContentBuilder
     $Builder | Add-BTText "Abriendo calendario"
     $Builder | Add-BTAppLogo "C:\Users\PHINX\Desktop\TFG\iconos\calendario.png"
     $Builder | Show-BTNotification
         Start-Sleep 1
            $posicion = python calendario.py
            $posicion = $posicion.replace("Box(","").replace(")","").split(",").trim()
                $MouseEvent = Add-Type -memberDefinition $MouseEventSig -name "MouseEventWinApi" -passThru
                [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(([Int]$posicion[0].replace("left=","")+15),([Int]$posicion[1].replace("top=","")+15))
                $MouseEvent::mouse_event(0x00000002, 0, 0, 0, 0)
                $MouseEvent::mouse_event(0x00000004, 0, 0, 0, 0)}


        if($orden.text -eq "cierra calendario") ##CIERRA EL CALENDARIO
    {
     $Builder = New-BTContentBuilder
     $Builder | Add-BTText "Cerrando calendario"
     $Builder | Add-BTAppLogo "C:\Users\PHINX\Desktop\TFG\iconos\calendario.png"
     $Builder | Show-BTNotification
        Start-Sleep 1
            Stop-Process -name HxCalendarAppImm}


    #MAXIMIZA#####################################################################################

        if($orden.text -eq "maximiza") ##MAXIMIZA LA VENTANA
    {$posicion = python maximiza.py
        $posicion = $posicion.replace("Box(","").replace(")","").split(",").trim()
        $MouseEvent = Add-Type -memberDefinition $MouseEventSig -name "MouseEventWinApi" -passThru
                  Start-Sleep -Seconds 1
                  [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(([Int]$posicion[0].replace("left=","")+15),([Int]$posicion[1].replace("top=","")+15))
                  $MouseEvent::mouse_event(0x00000002, 0, 0, 0, 0)
                  $MouseEvent::mouse_event(0x00000004, 0, 0, 0, 0)}

        if($orden.text -eq "minimiza") ##MINIMIZA LA VENTANA
    {$posicion = python minimiza.py
        $posicion = $posicion.replace("Box(","").replace(")","").split(",").trim()
        $MouseEvent = Add-Type -memberDefinition $MouseEventSig -name "MouseEventWinApi" -passThru
                  Start-Sleep -Seconds 1
                  [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(([Int]$posicion[0].replace("left=","")+15),([Int]$posicion[1].replace("top=","")+15))
                  $MouseEvent::mouse_event(0x00000002, 0, 0, 0, 0)
                  $MouseEvent::mouse_event(0x00000004, 0, 0, 0, 0)}

        if($orden.text -eq "cierra") ##CIERRA LA VENTANA
    {$posicion = python cierra.py
        $posicion = $posicion.replace("Box(","").replace(")","").split(",").trim()
        $MouseEvent = Add-Type -memberDefinition $MouseEventSig -name "MouseEventWinApi" -passThru
                  Start-Sleep -Seconds 1
                  [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(([Int]$posicion[0].replace("left=","")+15),([Int]$posicion[1].replace("top=","")+15))
                  $MouseEvent::mouse_event(0x00000002, 0, 0, 0, 0)
                  $MouseEvent::mouse_event(0x00000004, 0, 0, 0, 0)}


    #ALMACENAMIENTO###############################################################################OK

        if($orden.text -eq "disco duro") ##OBTIENE EL ALMACENAMIENTO RESTANTE DE LAS UNIDADES
    {
     $test = New-BTProgressBar -Status ‘_______________________________________________________’ -Indeterminate
        New-BurntToastNotification -ProgressBar $test -Text ‘Calculando espacio libre’ -AppLogo 'C:\Users\PHINX\Desktop\TFG\iconos\almacenamiento.ico' 
            Start-Sleep 2
                 foreach($inf in get-wmiobject -class win32_logicaldisk)
                    {
                    $esp = "Almacenamiento libre | "+[math]::Round(($inf).FreeSpace/1GB,2) +" GB"
                    $dat = "Alamcenamiento total | "+[math]::Round(($inf).size/1GB,2) +" GB"
                     $letra = "UNIDAD  " +($inf).DeviceID
                         New-BurntToastNotification -Text $letra,$dat,$esp -AppLogo 'C:\Users\PHINX\Desktop\TFG\iconos\dd.ico' 
                             start-sleep 3
                    }       
     }


    #EXPLORADOR###################################################################################

        if($orden.text -eq "explorador") ##ACCEDE EL EXPLORADOR
    {
     $Builder = New-BTContentBuilder
     $Builder | Add-BTText "Explorador de archivos"
     $Builder | Add-BTAppLogo "C:\Users\PHINX\Desktop\TFG\iconos\explorer.png"
     $Builder | Show-BTNotification
        Start-Sleep 1
            Start-Process explorer}


    #VMWARE#######################################################################################

	    if($orden.text -eq "maquina virtual") ##INICIA VMWARE	
    {
        $Builder = New-BTContentBuilder
        $Builder | Add-BTText "Administrador de maquinas virtuales","Espere un segundo"
        $Builder | Add-BTImage "C:\Users\PHINX\Desktop\TFG\iconos\vmware2.png"
        $Builder | Show-BTNotification
            Start-Sleep 2
                Start-Process "C:\Program Files (x86)\VMware\VMware Workstation\vmware.exe"}


        if($orden.text -eq "nueva") ##Parar
    {
        $posicion = python nuevamaquina.py
        $posicion = $posicion.replace("Box(","").replace(")","").split(",").trim()
        $MouseEvent = Add-Type -memberDefinition $MouseEventSig -name "MouseEventWinApi" -passThru
            Start-Sleep -Seconds 1
                [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(([Int]$posicion[0].replace("left=","")+15),([Int]$posicion[1].replace("top=","")+15))
                $MouseEvent::mouse_event(0x00000002, 0, 0, 0, 0)
                $MouseEvent::mouse_event(0x00000004, 0, 0, 0, 0)
    }


    #PANDELDECONTROL##############################################################################

	    if($orden.text -eq "panel de control") ##ABRE EL PANEL DE CONTROL	
    {
     $Builder = New-BTContentBuilder
     $Builder | Add-BTAppLogo "C:\Users\PHINX\Desktop\TFG\iconos\panel.png"
     $Builder | Add-BTText "PANEL DE CONTROL"
     $Builder | Show-BTNotification 
        Start-Sleep 1
            start ms-settings:}


    #CORTAFUEGOS##################################################################################

	    if($orden.text -eq "cortafuegos") ##ABRE EL CORTAFUEGOS	
    {
     $test = New-BTProgressBar -Status ‘_______________________________________________________’ -Indeterminate
        New-BurntToastNotification -ProgressBar $test -Text ‘Iniciando cortafuegos’ -AppLogo 'C:\Users\PHINX\Desktop\TFG\iconos\firewall.ico' 
            Start-Sleep 2
                Start-Process firewall.cpl}


        if($orden.text -eq "regla") ##ACCEDE A LAS REGLAS	
    {
        New-BurntToastNotification -Text ‘Reglas del cortafuegos’ -AppLogo 'C:\Users\PHINX\Desktop\TFG\iconos\firewall.ico' 
            Start-Sleep 2
                $posicion = python regla.py
                $posicion = $posicion.replace("Box(","").replace(")","").split(",").trim()
                $MouseEvent = Add-Type -memberDefinition $MouseEventSig -name "MouseEventWinApi" -passThru
                    Start-Sleep -Seconds 1
                        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(([Int]$posicion[0].replace("left=","")+15),([Int]$posicion[1].replace("top=","")+15))
                        $MouseEvent::mouse_event(0x00000002, 0, 0, 0, 0)
                        $MouseEvent::mouse_event(0x00000004, 0, 0, 0, 0)}


    #REGISTRO######################################################################################

	    if($orden.text -eq "registro") ##ABRE REGEDIT	
    {
     $Builder = New-BTContentBuilder
     $Builder | Add-BTAppLogo "C:\Users\PHINX\Desktop\TFG\iconos\registro.png"
     $Builder | Add-BTText "REGISTRO DE WINDOWS"
     $Builder | Show-BTNotification
        Start-Sleep 1
            Start-Process regedit}


    #ADAPTADORES####################################################################################

	    if($orden.text -eq "adaptador") ##INFORMACION ADAPTADORES	
    {
     foreach($ip in Get-NetIPConfiguration)
        {
         $Builder = New-BTContentBuilder
         $Builder | Add-BTAppLogo "C:\Users\PHINX\Desktop\TFG\iconos\adaptador.png"
         $interfaz = $ip.IPv4Address.InterfaceAlias
         $Builder | Add-BTText "$interfaz"
         $ip = $ip.IPv4Address.IPAddress
         $Builder | Add-BTText "Direccion IP: $ip"
         $enlace = $ip.IPv4DefaultGateway.NextHop
         $Builder | Add-BTText "Puerta de enlace: $enlace"
         $Builder | Show-BTNotification
        }
    }

    #ARBOL##########################################################################################

	    if($orden.Text -eq "arbol") ##EJECUTA UNA ESTRUCTURA EN ARBOL DEL PATH SELECCIONADO	
    {
     $Builder = New-BTContentBuilder
     $Builder | Add-BTText "Iniciando arbol"
     Add-BTHeroImage -ContentBuilder $Builder -Source 'C:\Users\PHINX\Desktop\TFG\iconos\arb.jpg'
     $Builder | Show-BTNotification
        Start-Sleep 2
            #Formulario
            $FormArbol = New-Object System.Windows.Forms.Form
            # Añadir una imagen al formulario
            $ImageArbol = [system.drawing.image]::FromFile("C:\Users\PHINX\Desktop\TFG\iconos\arbol.jpg")
            $FormArbol.BackgroundImage = $ImageArbol
            $FormArbol.Text= "Estructura en arbol"
            $FormArbol.Size=New-Object System.Drawing.Size(500,210)
            $FormArbol.StartPosition="CenterScreen"
            #Botón
            $ButtonArbol=New-Object System.Windows.Forms.Button
            $Button.Size=New-Object System.Drawing.Size(75,23)
            $ButtonArbol.Text="Crear"
            $ButtonArbol.Location=New-Object System.Drawing.Size(200,130)
            $ButtonArbol.add_click({tree $Var /F /a |Out-GridView;$FormArbol.Close()})
            $ButtonArbol.BackColor ="black"
            $ButtonArbol.ForeColor = "green"
            $ButtonArbol.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
            #Caja de texto
            $TextBoxArbol = New-Object System.Windows.Forms.TextBox
            $TextBoxArbol.Location = New-Object System.Drawing.Size(65,90)
            $TextBoxArbol.Size = New-Object System.Drawing.Size(350,20)
            $TextBoxArbol.BackColor = "black"
            $TextBoxArbol.ForeColor = "green"
            $TextBoxArbol.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 11, [System.Drawing.FontStyle]::Bold)
            #Funcionalidad para el formulario:
            #Pulsar la tecla Enter almacena en una variable el contenido de la caja de texto y se muestra
            #Pulsar la tecla Escape sale del formulario
            $FormArbol.KeyPreview = $True
            $FormArbol.Add_KeyDown({if ($_.KeyCode -eq "Enter"){$Var=$TextBoxArbol.Text;tree $var /f /a |Out-GridView;$FormArbol.Close()}})
            $FormArbol.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$FormArbol.Close()}})
            #Añadir botón
            $FormArbol.Controls.Add($ButtonArbol)
            #Añadir caja de texto
            $FormArbol.Controls.Add($TextBoxArbol)
            $FormArbol.ShowDialog()
    }

    #ANDEL##########################################################################################

        if($orden.text -eq "andel") ##Andel
    {
     $Builder = New-BTContentBuilder
     $Builder | Add-BTImage "C:\Users\PHINX\Desktop\TFG\iconos\andel.png"
     $Builder | Show-BTNotification
        
        start-sleep 1

        Add-Type -AssemblyName System.Speech
        gc C:\Users\PHINX\Desktop\TFG\andel\intro.txt | %{
        $synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
        $_
        #Establece la velocidad de habla de la SpeechSynthesizer
        $synthesizer.Rate=0
        $synthesizer.Speak($_)
    }


        [system.Diagnostics.Process]::Start("chrome","https://andel.es/quienes-somos/")
            Start-Sleep 3
        Add-Type -AssemblyName System.Speech
        gc C:\Users\PHINX\Desktop\TFG\andel\ANDEL.txt | %{
        $synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
        $_
        #Establece la velocidad de habla de la SpeechSynthesizer
        $synthesizer.Rate=0
        $synthesizer.Speak($_)
    }
           
        [system.Diagnostics.Process]::Start("chrome"," https://www.linkedin.com/in/pedro-luis-mu%C3%B1oz-8239b019/")
            Start-Sleep 3
        Add-Type -AssemblyName System.Speech
        gc C:\Users\PHINX\Desktop\TFG\\andel\pedro.txt | %{
        $synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
        $_
        #Establece la velocidad de habla de la SpeechSynthesizer
        $synthesizer.Rate=0
        $synthesizer.Speak($_)
    }
            
        [system.Diagnostics.Process]::Start("chrome"," https://www.linkedin.com/in/jesusninocamazon/")
            Start-Sleep 3
        Add-Type -AssemblyName System.Speech
        gc C:\Users\PHINX\Desktop\TFG\\andel\jesus.txt | %{
        $synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
        $_
        #Establece la velocidad de habla de la SpeechSynthesizer
        $synthesizer.Rate=0
        $synthesizer.Speak($_)
    }
            
        [system.Diagnostics.Process]::Start("chrome","https://www.linkedin.com/in/alfonsodeuna/")
            Start-Sleep 3
        Add-Type -AssemblyName System.Speech
        gc C:\Users\PHINX\Desktop\TFG\\andel\alfonso.txt | %{
        $synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
        $_
        #Establece la velocidad de habla de la SpeechSynthesizer
        $synthesizer.Rate=0
        $synthesizer.Speak($_)
    }

        [system.Diagnostics.Process]::Start("chrome","https://www.linkedin.com/in/ignacio-rico-teixeira/")
            Start-Sleep 3
        Add-Type -AssemblyName System.Speech
        gc C:\Users\PHINX\Desktop\TFG\\andel\ignacio.txt | %{
        $synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
        $_
        #Establece la velocidad de habla de la SpeechSynthesizer
        $synthesizer.Rate=0
        $synthesizer.Speak($_)
    }
            
        [system.Diagnostics.Process]::Start("chrome","https://www.linkedin.com/in/nicol%C3%A1s-dietl-881b951b1/")
            Start-Sleep 3
        Add-Type -AssemblyName System.Speech
        gc C:\Users\PHINX\Desktop\TFG\\andel\nicolas.txt | %{
        $synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
        $_
        #Establece la velocidad de habla de la SpeechSynthesizer
        $synthesizer.Rate=0
        $synthesizer.Speak($_)
    }
        Add-Type -AssemblyName System.Speech
        gc C:\Users\PHINX\Desktop\TFG\\andel\despedida.txt | %{
        $synthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
        $_
        #Establece la velocidad de habla de la SpeechSynthesizer
        $synthesizer.Rate=0
        $synthesizer.Speak($_)
    }
    }

    #STOP###########################################################################################

        if($orden.text -eq "parar") ##Parar
    {
        $Builder = New-BTContentBuilder
        $Builder | Add-BTText "Hasta luego!"
        Add-BTHeroImage -ContentBuilder $Builder -Source 'C:\Users\PHINX\Desktop\TFG\iconos\cat3.gif'
        $Builder | Show-BTNotification
            break
    }

    
    #ORDENES########################################################################################

        if($orden.text -eq "aceptar") ##Parar
    {
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}");start-sleep 1

    }

    ###

        if($orden.text -eq "tabular") ##Parar
    {
        [System.Windows.Forms.SendKeys]::SendWait("{TAB}");start-sleep 1

    }

    ###

        if($orden.text -eq "escape") ##Parar
    {
        [System.Windows.Forms.SendKeys]::SendWait("{ESC}");start-sleep 1

    }

    ###

        if($orden.text -eq "arriba") ##Parar
    {
        [System.Windows.Forms.SendKeys]::SendWait("{UP}");start-sleep 1

    }

    ###

        if($orden.text -eq "abajo") ##Parar
    {
        [System.Windows.Forms.SendKeys]::SendWait("{DOWN}");start-sleep 1

    }
    
    ###

        if($orden.text -eq "izquierda") ##Parar
    {
        [System.Windows.Forms.SendKeys]::SendWait("{LEFT}");start-sleep 1

    }
    
    ###

        if($orden.text -eq "derecha") ##Parar
    {
        [System.Windows.Forms.SendKeys]::SendWait("{RIGHT}");start-sleep 1

    }

}while(1)
