#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.4.1
 Author:         Sven Neuhaus <sven-theta AT sven DOT de>
 Script Function:
	Ricoh Theta 360° image batch processing
	Copy to directory with images taken by a Ricoh Theta and run it.
	It will generate auto-rotated images with XMP metadata.

    Will not work properly if there are more than 9999 files that have
	already	been processed in the current directory.

#ce ----------------------------------------------------------------------------

; Config Section ---------------------------------------------------------------

; Ensure path to 'RICOH THETA.exe' is correct for your system
;Local $theta_exe = @ProgramFilesDir & "\RICOH THETA\RICOH THETA.exe"
Local $theta_exe = "C:\Program Files (x86)\RICOH THETA\RICOH THETA.exe"

; set doVideo to 1 to perform MP4 spherical video conversion
; the default is 0, perform photo conversion
; warning: the video option currently does some rather wild clicking
Local $doVideo = 1

; Use these for German Locale
;Local $theta_title = "RICOH THETA"
;Local $theta_save_title = "JPEG-Daten mit XMP"
;Local $shortcut_file_menu = "!D"
;Local $shortcut_write_with_up_down = "m"
;Local $shortcut_write_xmp = "j"
;Local $convertVideoTitle = "Video Konvertieren" 	; a guess

; Use these for UK Locale
Local $theta_title = "RICOH THETA"
Local $theta_save_title = "JPEG data with XMP"
Local $shortcut_file_menu = "!F"
Local $shortcut_write_with_up_down = "w"
Local $shortcut_write_xmp = "j"
Local $convertVideoTitle = "Convert Video"

; No user serviceable parts below ----------------------------------------------

Local $fileExt = ".JPG";
Local $convertedFileExt = "_xmp.JPG";
Local $wildcard = "*";
Local $sleepTime = 1000  ; 1 second

If $doVideo = 1 Then
   $fileExt = ".MP4";
   $convertedFileExt = "_er.MP4";
   $sleepTime = 5000	; 5 seconds. video conversion tends to take ages
EndIf


#include <File.au3>
Local $donefiles[9999] ; it's over 9000!
$donefiles = preprocessed_files()

; Now we go through all files with the required extension in the directory
Local $search_handle = FileFindFirstFile( $wildcard & $fileExt)
If $search_handle = -1 Then
    MsgBox($MB_SYSTEMMODAL, "", "Error: No " & $fileExt & " files in current directory.")
 EndIf

AutoItSetOption ("PixelCoordMode", 0)  ;relative to Window

Local $done = 0
Do
	Do
		$file = FileFindNextFile($search_handle)
		If @error Then
			$done = 1
			ExitLoop
		EndIf
	Until Not(StringInStr($file, $convertedFileExt))

	If ($done = 1) Then
		ExitLoop
	EndIf

	; check for file in list of already completed images ($donefiles)
	Local $dfile
	For $dfile In $donefiles
		Local $skip = 0
		if (StringInStr($file, $dfile) = 1) Then
			ConsoleWrite("skipping '" & $file & "' (" & $convertedFileExt & " file already present)" & @CRLF)
			$skip = 1;
			ExitLoop
		EndIf
	Next

	if ($skip <> 1) Then
	  if $doVideo = 1 Then
		 processVideo($file)
	  Else
		xmp_image($file)
	  EndIf
	EndIf

Until $done = 1

FileClose($search_handle)
MsgBox($MB_APPLMODAL, "", "The autoit script is finished.")


; get a list of all already processed files (*_xmp.JPG / *_er.MP4)
Func preprocessed_files ()
	Local $file
	Local $donefiles[9999] ;"WHAT?! NINE THOUSAND?!"
	Local $di = 0
	Local $search_handle = FileFindFirstFile( $wildcard & $convertedFileExt )
	If $search_handle <> -1 Then
		Do
			$file = FileFindNextFile($search_handle)
			If @Error Then
				ExitLoop
			EndIf
			; remember the filename but without the "_xmp.JPG" suffix
			$donefiles[$di] = StringLeft($file, StringInStr( $file, $convertedFileExt )-1)
			$di = $di + 1
		Until @error
		FileClose($search_handle)
	EndIf
	Return $donefiles
EndFunc


; open file with theta app then save it with xmp data
Func xmp_image($image_file)
	; open program with image as parameter
	Run($theta_exe & ' "' & @WorkingDir & '\' & $image_file & '"')
	WinWaitActive($theta_title)
	; after loading the image, the window title changes
	WinWaitActive($image_file & " - " & $theta_title)
	; image loading done.

	; Save the file with XMP data and rotation correction

;	WinMenuSelectItem($image_file & " - " & $theta_title, "", "&Datei", "&Mit oben/unten schreiben", "&JPEG-Daten mit XMP")
	; File Menu ("Datei")
	Send($shortcut_file_menu)
	; Submenu "Mit oben/unten schreiben"
	Send($shortcut_write_with_up_down)
	; Menuentry "JPEG-Daten mit XMP"
	Send($shortcut_write_xmp)
	; wait for "save file" dialog to open
	WinWaitActive($theta_save_title)
	Send("{ENTER}")

	; Wait until the file has been written
	wait_until_ready($image_file)

	WinClose($image_file & " - " & $theta_title)
EndFunc

; open file with theta app then save it with xmp data
Func processVideo($video_file)
	; open program with image as parameter
	Run($theta_exe & ' "' & @WorkingDir & '\' & $video_file & '"')
	WinWaitActive($theta_title)

	; after loading the video, the 'Convert Video' window is displayed automatically
	WinWaitActive($convertVideoTitle)

   ; for videos the converted file name is used in the title once conversion is complete
   Local $convertedName = StringReplace( $video_file, $fileExt, $convertedFileExt, -1 );

	; we'll accept the defaults in 'Convert Video',
	; i.e. same to same dir, with _er.MP4 extension
	; TBD
	; Issue clicking 'Start';
	; I can't find the control id and 5 tabs/return fails, and as my knowledge
	; of autoit is limited i'm going for the nuts approach of clicking wildly
	; bottom left to bottom right in a Start button scan of madness.
	; On my machine the Start button is at $width-135, $height-40, but we'll
	; do a sweep of the bottom right of the window.

   ; calculate scan box
   ; (bottom right of window, scanning direction is left to right, bottom to top)
   Opt("MouseCoordMode", 0) ; coords relative to active window
   Local $winSize = WinGetClientSize( $convertVideoTitle )
   Local $width = $winSize[0]
   Local $height = $winSize[1]

   Local $yScanStart = $height-20
   Local $yScanEnd = $height/2
   Local $xScanStart = $width/2
   Local $xScanEnd = $width - 50;

   ; do scan (to click start button and commence conversion)
   Local $success = 0

   For $yScan = $yScanStart To $yScanEnd Step -15

	  For $xScan = $xScanStart To $xScanEnd Step 30

		 MouseClick("left", $xScan, $yScan, 1, 1 )

		 Local $currentName = WinGetTitle("[ACTIVE]");

		 If $currentName <> $convertVideoTitle Then
			If $currentName <> $convertedName And $currentName <> $theta_title And $currentName <> "" Then
			   MsgBox($MB_SYSTEMMODAL, "", "Error: unexpect window " & $currentName)
			Else
			   ; window has gone, with any luck we managed to click Start
			   $success = 1;
			Endif
			ExitLoop 2 ; we've either hit an error or succeeded
		 EndIf
	  Next
   Next

   If $success = 1 Then
	  ; Wait until the file has been written
	  wait_until_ready($video_file)

	  WinClose($convertedName & " - " & $theta_title)
   Else
	  MsgBox($MB_SYSTEMMODAL, "", "Aborting conversion")
	  WinClose($theta_title)
   EndIf
EndFunc

; This function waits until the program has finished saving the image
Func wait_until_ready($file)

   $target_file = StringLeft($file, StringInStr( $file, $fileExt,-1)-1) & $convertedFileExt

   While (Not FileExists($target_file))
	  Sleep($sleepTime)
   WEnd

#comments-start
	Local $handle = WinGetHandle($image_file & " - " & $theta_title)
	; check if menu is enabled (doesn't work)
	;ControlCommand($hWin, "", "[NAME:button2]", "IsEnabled", "")
    $pixelx = 473; // was 914
	$pixely = 735; // was 930
	; the "+" button will be grayed out, then it will turn white again
	$color = PixelGetColor($pixelx, $pixely, $handle )
	While $color = 0xFFFFFF
		Sleep(50)
		$color = PixelGetColor($pixelx, $pixely, $handle )
		; TODO: if computer is too fast it may have been grey before we saw it... XXX
	WEnd

	ConsoleWrite(" pixel is no longer white" & @CRLF)
	$color = PixelGetColor($pixelx, $pixely, $handle)
	While $color <> 0xFFFFFF
		Sleep(100)
		$color = PixelGetColor($pixelx, $pixely, $handle)
	WEnd

	ConsoleWrite(" pixel is white again, writing has finished." & @CRLF)
#comments-end

EndFunc

;eof. This file has not been truncate
