   #Information prompt

   $wshell = New-Object -ComObject Wscript.Shell
   $exceptionMessage = 
"This program will move a selected downloaded certificate to the same folder as the program. (This is should be a USB or a certificate backup folder). It will then start the certloader program to install your HealthSecure certificate.
   
The downloaded certificate will be cert.p12 which is renamed to the required filename cert_<ddmmyy>.pfx."
 
   $wshell.Popup($exceptionMessage, 0, "NZ Ministry of Health HealthSecure Certificate Install")
 
    # Get the current directory of the PowerShell script
    $currentDirectory = Get-Location

    # Create a file select dialog box
    Add-Type -AssemblyName System.Windows.Forms
    $fileSelectDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileSelectDialog.Title = "Select the your healthsecure download file (cert)"
    $fileSelectDialog.Filter = "All files (cert*.*)|cert*.*"

    # Display the file select dialog box and get the selected file path
    $fileSelectDialogResult = $fileSelectDialog.ShowDialog()
    $selectedFilePath = $fileSelectDialog.FileName

    # If the user selected a file, copy and rename it
    if ($selectedFilePath) {

        # Get the current date and time
        $currentDate = Get-Date -Format "ddMMyy"

        # Define the new file name
        $newFileName = "cert_$currentDate.pfx"

    # Check if the destination file already exists
    if (Test-Path "$currentDirectory\$newFileName") {

        # Increment the file name until it is unique
        $i = 1
        while (Test-Path "$currentDirectory\($i)$newFileName") {
            $i++
        }

        # Rename the file to make it unique
        $newFileName = "($i)$newFileName"
    }

         try {

        # Rename the file
        #Rename-Item -Path $selectedFilePath -NewName $newFileName
        Move-Item -Path $selectedFilePath -Destination $currentDirectory\$newFileName
         }

        catch {
            $wshell = New-Object -ComObject Wscript.Shell
            $exceptionMessage = "Failed to rename this file " + $selectedFilePath + " to " +$currentDirectory +"\"+ $newFileName +". Error message : " + $_.Exception.Message
            $wshell.Popup($exceptionMessage, 0, "Error occured")
        }

        # Get the path to the CertLoader-setup.exe program
        $certLoaderSetupExePath = "CertLoader-setup.exe"

        # Start the CertLoader-setup.exe program
        Start-Process "$certLoaderSetupExePath"
    }

    
 