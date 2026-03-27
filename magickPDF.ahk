#Requires AutoHotkey v2.0

global OutFolder := ""
global SessionFolder := ""
global CapX := 0, CapY := 0, CapW := 0, CapH := 0
global SnapCount := 0

; Main GUI

MainGui := Gui("AlwaysOnTop", "magickPDF")
MainGui.SetFont("s10", "Segoe UI")

MainGui.Add("Text", "w300", "1. Select where to save the images/PDF:")
TxtFolder := MainGui.Add("Text", "w300 r2 cGray", "No folder selected.")
BtnFolder := MainGui.Add("Button", "w300", "Select Destination")
BtnFolder.OnEvent("Click", ChooseFolder)

MainGui.Add("Text", "w300 y+20", "2. Define capture area:")
BtnArea := MainGui.Add("Button", "w300", "Select Dimensions")
BtnArea.OnEvent("Click", SetupArea)

MainGui.Add("Text", "w300 y+20", "3. Name your PDF (optional):")
TxtPDFName := MainGui.Add("Edit", "w300", "Final_Document")

TxtStatus := MainGui.Add("Text", "w300 y+20 cBlue", "Status: Waiting for setup...")

MainGui.OnEvent("Close", (*) => ExitApp())
MainGui.Show()


; Core Functions

ChooseFolder(*) {
    global OutFolder, SessionFolder
    selected := DirSelect(A_ScriptDir, 3, "Select Base Folder")
    if (selected != "") {
        OutFolder := selected
        SessionFolder := OutFolder "\PDF_Snaps_" FormatTime(, "yyyyMMdd_HHmmss")
        try {
            DirCreate(SessionFolder)
            TxtFolder.Value := SessionFolder
            TxtStatus.Value := "Status: Folder created. Now select dimensions."
        } catch Error as e {
            MsgBox("Could not create folder: " e.Message)
        }
    }
}

SetupArea(*) {
    global BoxGui
    BoxGui := Gui("+Resize +AlwaysOnTop -Caption +Border +ToolWindow", "Resize Me")
    BoxGui.BackColor := "Red"
    WinSetTransparent(100, BoxGui)
    
    btn := BoxGui.Add("Button", "x10 y10 w150 h40", "Confirm Dimensions")
    btn.OnEvent("Click", SaveDimensions)
    
    BoxGui.Show("w600 h800 x100 y100")
    TxtStatus.Value := "Status: Resize the red box, then click Confirm."
}

SaveDimensions(*) {
    global CapX, CapY, CapW, CapH, BoxGui
    WinGetPos(&CapX, &CapY, &CapW, &CapH, BoxGui.Hwnd)
    BoxGui.Destroy()
    TxtStatus.Value := "READY!`nF8: Capture Page`nF9: Compile & Clean"
}

; Keybinds
; f8 to trigger a screenshot
; f9 to compile all screenshots into a single pdf
F8:: {
    global SnapCount, SessionFolder, CapX, CapY, CapW, CapH
    if (SessionFolder = "" or CapW = 0) {
        ToolTip("Error: Configure the folder and dimensions first.")
        SetTimer(() => ToolTip(), -2000)
        return
    }
    
    SnapCount++
    fileName := SessionFolder "\page_" Format("{:03}", SnapCount) ".png"
    
    cmd := 'magick screenshot: -crop ' CapW 'x' CapH '+' CapX '+' CapY ' +repage "' fileName '"'
    
    try {
        RunWait(A_ComSpec ' /c ' cmd, , "Hide")
        ToolTip("Captured Page " SnapCount)
    } catch {
        ToolTip("Error: Is ImageMagick installed and in PATH?")
    }
    
    SetTimer(() => ToolTip(), -1000)
}

F9:: {
    global SessionFolder, SnapCount, OutFolder
    if (SessionFolder = "" or SnapCount = 0) {
        MsgBox("No images captured yet.")
        return
    }
    
    pdfName := TxtPDFName.Value
    if (pdfName = "")
        pdfName := "Final_Document"
    
    if (SubStr(pdfName, -3) = ".pdf")
        pdfName := SubStr(pdfName, 1, StrLen(pdfName) - 4)

    TxtStatus.Value := "Status: Compiling PDF... (Please wait)"
    finalPDF := SessionFolder "\" pdfName ".pdf"
    
    cmd := 'magick "' SessionFolder '\*.png" "' finalPDF '"'
    
    try {
        RunWait(A_ComSpec ' /c ' cmd, , "Hide")
        
        SourceImagesDir := SessionFolder "\_source_images"
        DirCreate(SourceImagesDir)
        
        Loop Files, SessionFolder "\*.png" {
            FileMove(A_LoopFileFullPath, SourceImagesDir "\" A_LoopFileName, 1)
        }

        ; Post Compile, we rename the project folder to match the name of the PDF in question
        renamedFolder := OutFolder "\" pdfName
        DirMove(SessionFolder, renamedFolder, 2)
        SessionFolder := renamedFolder

        TxtStatus.Value := "Status: Finished! PDF created and images archived."
        Run('explorer.exe "' SessionFolder '"')
        
    } catch Error as e {
        MsgBox("Compilation failed: " e.Message)
    }
    
    SnapCount := 0
}
