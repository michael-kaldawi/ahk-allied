#Requires AutoHotkey v2.0
#SingleInstance Force
#include FindText.ahk

; Used for testing purposes.
; F1:: 
; {
;     MsgBox "The active window is '" WinGetTitle("A") "'. \nThe active class is '" WinGetClass("A") "' \nThe active window's text is '" WinGetText("A") "' ."
; }

; global spreadsheet title
global BADGE_SPREADSHEET := "Seattle badges distribution log"

; global badge text image (in text form)
global BADGE_TEXT := Text:="|<>*143$427.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzVzzzzzzzzzzzzzzzzVzzzzzzsTzzzzzzwDzzzzzzzzzzkzzzzzzzzzzzzzzzzzzzzzzzzzzkzzzzzzzzzzzzzzzzkzzzzzzwDzzzzzzy7zzzzzzzzzzsTzzzzzzzzzzzzzzzzzzzzzzzzzsTzzzzzzzzzzzzwTwMTzzzzzy7zzzzzzz3zzzzzzzzzzwDzzzzzzzzzzzzzzzzzzzzzzwDzwDzzzzzzzzzzzzwDwADzzzzzz3zzzzzzzVzzzzzzzzzzy7zzzzzzzzzzzzzzzzzzzzzzyDzy7zz7yDzzzzzzzyDy67zzzzzzVzzzzzzzkzzzzzXzzzzz3zzzzzzzzzzzzzzzzzzzzzzy7zz3zz3y7zzzzzzzy7z73zzzzzzkzzzzzzzsTzzzzVzzzzzVzzzzzzzzzzzzzzzzzzzzzzz7zzVzzVz3zzzzzzzz7z3VzzzzzzsTzzzzzzwDzzzzkzzzzzkzzzzzzzzzzzzzzzzzzzzzzz3zzkzzkzVzzzzzzzzXzXkzzzzzzwDzzzzzzy7zzzzsTzzzzsTzzzzzzzzzzzzzzzzzzzzzzVzzsMDU10330zk7VzVzVsM7y0zs67w67w1zz30wDs00y0zs4Dk7wA7zUDs7zzk7w1zX1s7zlwTw81k0U1V0Dk1kzlzlw81w07s13s03s0TzV0C3w80S0Dk07U0y01z07k1zzU1s0Dk081zky/y00s0E0k03k0sTkzky00S03s01s01w07zk03VyA0C03s03k0T00T03k0TzU0s03s000Tsz1z1sD3y7s7UkzwDszsT0w7DksD0wDUwD3zsDVky7Vy7ksD1ty7UwD1xky7zkSQDVw70wDsTUzUy7Vz3w7sMTzzsTwTUz3zsQDkQDsQDkzw7sQD3kz7sQDkzz3ky7VzsT3zkzwDsS7ky7wTkzkz3kzVy7wC7zzwDwDkzVzwC7wC7wC7sTy7wC7XsT3wADsTzVsTXkzwTlzsTy7wD3sT3wDsTsTVsTkz3y70TzyDyDsTkz077y73y700Dz3y73VwDU067wDs0wDkkzw00zwTz7y7XwDly7wTwDkwDsTVz3k3zy7y7wDsS03Xz3Vz3U07zVz3kky7k033y7k0S7sMTy00TyDzXzXly7sz7yDy7sS7wDkzVy0zz7z3y7wC01lzVkzVk03zkzVsMz3s01Vz3k0D3wADz00Dz7zlzVsz3wT3z7z3wD3y7sTkzsTz3zXz3y67ksTksTksTzzsTky8TVwDzkzVky7Vy63zVzzzVzsTkwTVyDXzXzVy7Vz3w7sTwDzXzVzUz33sQDkQDkQDzzwDkz0Tky7zsDksT3kz3VzsTzzkzwDsSDkz7Vzkzkz3sTVy3sTy73VzlzkD3VsC3kC3kD3wwS3kTUDsTVzS7kQD1sTVkTQ7vksDa3sT7sTXlzsDsTVw1k300C03Vkzkzs01k07U07U07U0Q700Ts7w0k0D00C00wDkw0D01sS03U0DXwDlszw7wDky0w1V0D03kszszw81w03s13s13s0C3U0Dw7z0Q07k07U0S7sT07k0wDU1s0Dly7ssTy3y7sTUT0kkTk3sMTsTy63z0Vy1Vy1Vz0TXkUTz3zkDU7w33s4D3wDk3w0y7w1z0Tsz3wQTz7zzzzzzzsTzzzzwTwDzzzzzzzzzzzkzzzzzzzzXzzzzzzzzzzzzzzzzzzzzzzzzzzzzzwDzXzzzzzzzwDzzzzwDyDzzzzzzzzzzzsTzzzzzzzVzzzzzzzzzzzzzzzzzzzzzzzzzzzzzyDzlzzzzzzzy7zzzzyDy7zzzzzzzzzzzsTzzzzzzzkzzzzzzzzzzzzzzzzzzzzzzzzzzzzzy7zszzzzzzzz3zzzzzzzzzzzzzzzzzzbsDzzzzzzzkzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzwTzzzzzzzVzzzzzzzzzzzzzzzzzzk0Dzzzzzzz0zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzyDzzzzzzzkzzzzzzzzzzzzzzzzzzs0DzzzzzzzUTzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz7zzzzzzzsTzzzzzzzzzzzzzzzzzy0TzzzzzzzkzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzXzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzw"

; open these apps upon startup
OpenBadgeSheet()
OpenCamera()
OpenChrome()

; poll for badge links on-screen every 50ms
SetTimer DetectBadgeLink, 50

DetectBadgeLink()
{
    X:=Y:=""

    Text:=BADGE_TEXT

    if (ok:=FindText(&X, &Y, 354-150000, 871-150000, 354+150000, 871+150000, 0, 0, Text))
    {
    FindText().Click(X, Y, "L")
    Sleep 600
    ; select address bar
    Send "^{l}"
    Sleep 100
    ; copy content
    Send "^{c}"
    Sleep 100
    ; close the opened chrome tab
    Send "^{w}"
    Sleep 100

    ; parse contents of clipboard; copy the badge number
    FoundPos := InStr(A_Clipboard, "=")
    badgeNum := SubStr(A_Clipboard, FoundPos+1, 9)
    ; MsgBox "The badge number is:`n" badgeNum
    A_Clipboard := badgeNum
    
    ; minimize chrome and open the spreadsheet
    WinMinimize("Chrome")
    WinMaximize(BADGE_SPREADSHEET)
    Sleep 600

    ; paste the badge number into the first open row of the first column
    Click 1000, 600
    Sleep 100
    Send "^{Home}"
    Sleep 100
    Send "^{Down}"
    Sleep 100
    Send "{Down}"
    Sleep 100
    Send "^{v}"
    Sleep 100

    ; enter the date
    currentDate := FormatTime(, "MM/dd/yyyy")
    Send "{Right}"
    Sleep 100
    Send currentDate
    Sleep 100
    Send "{Enter}"

    ; WinActivate "Camera"
    Sleep 10000
    }

}

OpenBadgeSheet()
{
    ; open the web app
    Run 'cmd.exe /c ""C:\Program Files\Google\Chrome\Application\chrome_proxy.exe"  --profile-directory="Profile 2" --app-id=cnonfnndahheilompemimnjomgecihol"'
    WinWait(BADGE_SPREADSHEET)
    WinMinimize(BADGE_SPREADSHEET)
}

OpenCamera()
{
    Run "cmd.exe /c start microsoft.windows.camera:"
    WinWait "Camera"
    Send "{LWin down} {Left}"
    Send "{LWin up}"
}

OpenChrome()
{
    Run 'C:\Program Files\Google\Chrome\Application\chrome.exe'
    WinWait "Chrome"
    WinMinimize("Chrome")
}
