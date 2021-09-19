Attribute VB_Name = "m00_ver_control"
'Option Explicit

Sub dodawanie_referencji()
'programowe dodawanie referencji do Microsoft Visual Basic Extensibility 5.3
'------------------------------------------------------------------------------------------------------------------------------------------------------
'    TEGO MAKRA razem NIE DA SIE WYKONAC KROKOWO (tylko F5)
'    jezeli udaje Ci sie uruchomic to makro krokowo to pewnie nie masz wlaczonej opcji "Ufaj dostepowi do modelu obiektowego Projektu VBA"
'------------------------------------------------------------------------------------------------------------------------------------------------------

On Error Resume Next
   ThisWorkbook.VBProject.References.AddFromGuid GUID:="{0002E157-0000-0000-C000-000000000046}", Major:=5, Minor:=3

End Sub

Sub eksportuj_caly_kod_z_pliku()

'dodaj refeerencje do VBE Extensibility
    Call dodawanie_referencji

'sprawdzam czy uzytkownik ma wlaczony w bezpieczenstwie makr opcje "Ufaj dostepowi do modelu obiektowego Projektu VBA"
    If Not czy_jest_dostep_do_modelu_obiekt_VBA Then
       MsgBox "Zeby kod dzialal, wlacz najpierw w bezpieczenstwie makr opcje:" & Chr(10) & Chr(10) & _
              """Ufaj dostepowi do modelu obiektowego Projektu VBA""", vbCritical
       Exit Sub
    End If
    
'sprawdzam czy Uzytkownik ma otwarte okno z edytorem VBA
    If (ThisWorkbook.VBProject.VBE.ActiveWindow Is Nothing) Then
        Exit Sub
    End If
    
'wyciagnij sciezke do eksportowania ze sciezki do pliku
    Dim sciezka_eksport As String
    Dim tabik() As String

    tabik = Split(ThisWorkbook.fullName, "\")
        ReDim Preserve tabik(UBound(tabik) - 1)
        sciezka_eksport = Join(tabik, "\") & "\eksporcik\"
    
'stworz katalog na eksportowane pliki
    On Error Resume Next
        MkDir sciezka_eksport
    On Error GoTo 0
    
'zrob eksport
    Dim komponent_kodu 'As VBComponent
    Dim do_eksportu As String
    Dim timestamp As String
    
    timestamp = Format(Now, "_yyyymmdd_hhmmss")
    
On Error Resume Next
    For Each komponent_kodu In ThisWorkbook.VBProject.VBComponents
        Select Case komponent_kodu.Type
            Case 1 '  1-->vbext_ct_StdModule
                MkDir sciezka_eksport & "modules\"
                do_eksportu = sciezka_eksport & "modules\" & komponent_kodu.Name & timestamp
                komponent_kodu.Export do_eksportu & ".bas"
            Case 2 '  2-->vbext_ct_ClassModule
                MkDir sciezka_eksport & "classes\"
                do_eksportu = sciezka_eksport & "classes\" & komponent_kodu.Name & timestamp
                komponent_kodu.Export do_eksportu & ".cls"
            Case 3 '  3-->vbext_ct_MSForm
                MkDir sciezka_eksport & "forms\"
                do_eksportu = sciezka_eksport & "forms\" & komponent_kodu.Name & timestamp
                komponent_kodu.Export do_eksportu & ".frm"
        End Select
    Next
'test

End Sub

Sub zrob_commita()
    
    Shell "git commit -m 'update ' & format(now, ""yyyymmdd_hhmmss"")"
    
End Sub

Function czy_jest_dostep_do_modelu_obiekt_VBA() As Boolean
    Dim wsh
    Dim klucz As String
    Dim wartosc_klucza As Long

'tworze nowa instancje obiektu Wscript
    Set wsh = CreateObject("WScript.Shell")
    
'sklejam do kupy klucz rejestru w ktorym Windows przechowuje ustawienie o dostepie do modelu obiektowego
    klucz = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Excel\Security\AccessVBOM"

'sprawdzamy wartosc w rejestrze Windowsa
    On Error Resume Next
        wartosc_klucza = wsh.RegRead(klucz)
    On Error GoTo 0
    
'zamiast IFa robie porownanie czy wartosc klucza=1 jak tak to funkcja bedzie miaal wartosc TRUE a jak nie to False
    czy_jest_dostep_do_modelu_obiekt_VBA = (wartosc_klucza = 1)
   
'sprzatanie
    Set wsh = Nothing
End Function
