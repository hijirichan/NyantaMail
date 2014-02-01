#ifndef STRICT
    #define STRICT
#endif

#ifndef WIN32_LEAN_AND_MEAN
    #define WIN32_LEAN_AND_MEAN
#endif

#define ID_EDIT 100
#define ID_STATUS 101
#define SZMYVERSION "にゃんたメールのﾊﾞｰｼﾞｮﾝ情報#にゃんたメール Version 0.1"
#define SZMYSZMYCOPYRIGHT "Copyright (C) 2002 砂倉瑞希"

#include <windows.h>
#include <winsock2.h>
#include <windowsx.h>
#include <commctrl.h>
#include <shellapi.h>	// ShellAbout()を使うのに必要
#include <stdio.h>
#include <stdlib.h>
#include <mbstring.h>
#include "resource.h"

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK MyDlgProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK MySettingProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK MyPopSetProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK MyNewMailProc(HWND, UINT, WPARAM, LPARAM);

ATOM InitApp(HINSTANCE);
BOOL InitInstance(HINSTANCE, int);

void MySnd(HWND, char *, char *, char *);
void MyRcv(HWND, HWND, HWND, HWND);
int GetMailSize(char *);
void MyJisToSJis(char *, char *);
void MySJisToJis(char *, char *);
void GetSubject(char *, char *);
void GetFrom(char *, char *);
void GetDate(char *, char *);
void GetString(char *, char *);
void decode(char *, char *);
char charconv(char);
void InsertColumn(HWND);
void InsertItem(HWND, int, int, char *);
void ShowContents(HWND, HWND, HWND, int);
void MyPopConnect(HWND, HWND);
void MyPopDisconnect(HWND);
void encode(char *, char *);
char convtobase(char);

char szClassName[] = "nmail";			// ウィンドウクラス
char szStr[1024], szStrRcv[1024 * 50], szResult[1024 * 50];
char szServerName[256], szFrom[256], szReplyTo[256];	// SMTP
char szPopServer[256], szUserName[64], szPass[64];		// POP3

HINSTANCE hInst;

SOCKET s;			// ソケット(POP3用)
BOOL bPOP = NULL;	// POP3接続中かどうか

int WINAPI WinMain(HINSTANCE hCurInst, HINSTANCE hPrevInst, LPSTR lpsCmdLine, int nCmdShow)
{
	MSG msg;
	
	hInst = hCurInst;
	
	if(!InitApp(hCurInst))
		return FALSE;
	
	if(!InitInstance(hCurInst, nCmdShow))
		return FALSE;
	
	while(GetMessage(&msg, NULL, 0, 0)){
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return msg.wParam;
}

// ウィンドウ・クラスの登録
ATOM InitApp(HINSTANCE hInst)
{
	WNDCLASSEX wc;
	wc.cbSize = sizeof(WNDCLASSEX);
	wc.style = CS_HREDRAW | CS_VREDRAW;
	wc.lpfnWndProc = WndProc;			// プロシージャ名
	wc.cbClsExtra = 0;
	wc.cbWndExtra = 0;
	wc.hInstance = hInst;				// インスタンス
	wc.hIcon = LoadIcon(hInst, "MYICON");
	wc.hCursor = LoadCursor(NULL, IDC_ARROW);
	wc.hbrBackground = (HBRUSH)GetStockObject(WHITE_BRUSH);
	wc.lpszMenuName = "MYMENU";    //メニュー名
	wc.lpszClassName = (LPCSTR)szClassName;
	wc.hIconSm = LoadIcon(hInst, "MYICON");
	return (RegisterClassEx(&wc));
}

// ウィンドウの生成
BOOL InitInstance(HINSTANCE hInst, int nCmdShow)
{
	HWND hWnd;
	
	hWnd = CreateWindow(szClassName,
		"にゃんたメール",		//タイトルバーにこの名前が表示されます
		WS_OVERLAPPEDWINDOW,	//ウィンドウの種類
		CW_USEDEFAULT,			//Ｘ座標
		CW_USEDEFAULT,			//Ｙ座標
		CW_USEDEFAULT,			//幅
		CW_USEDEFAULT,			//高さ
		NULL,					//親ウィンドウのハンドル、親を作るときはNULL
		NULL,					//メニューハンドル、クラスメニューを使うときはNULL
		hInst,					//インスタンスハンドル
		NULL);
	
	if(!hWnd)
		return FALSE;
	
	ShowWindow(hWnd, nCmdShow);
	UpdateWindow(hWnd);
	
	return TRUE;
}

//ウィンドウプロシージャ
LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wp, LPARAM lp)
{
    int id, iClickNo;
    static HWND hEdit, hStatus, hList;
    INITCOMMONCONTROLSEX ic;
    static int iStatusWy;
    RECT rc;
    DWORD dwStyle;
    LPNMHDR lpnmhdr;
    LPNMLISTVIEW lpnmlv;
	HICON hIcon;

    switch (msg) {
        case WM_CREATE:
            ic.dwSize = sizeof(INITCOMMONCONTROLSEX);
            ic.dwICC = ICC_BAR_CLASSES | ICC_LISTVIEW_CLASSES;
            InitCommonControlsEx(&ic);
            hEdit = CreateWindowEx(0,
                "EDIT", "",
                WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_WANTRETURN |
                WS_VSCROLL | WS_HSCROLL | ES_AUTOVSCROLL | ES_AUTOHSCROLL,
                0, 0, 0, 0, hWnd, (HMENU)ID_EDIT, hInst, NULL);
            hList = CreateWindowEx(0, WC_LISTVIEW, "",
                WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT,
                0, 0, 0, 0, hWnd, (HMENU)100, hInst, NULL);
            dwStyle = ListView_GetExtendedListViewStyle(hList);
            dwStyle |= LVS_EX_FULLROWSELECT |
                LVS_EX_GRIDLINES | LVS_EX_HEADERDRAGDROP;
            ListView_SetExtendedListViewStyle(hList, dwStyle);
            InsertColumn(hList);
    
            hStatus = CreateStatusWindow(WS_CHILD | WS_VISIBLE, "", hWnd, ID_STATUS);
            GetWindowRect(hStatus, &rc);
            iStatusWy = rc.bottom - rc.top;
            break;
        case WM_SIZE:
            MoveWindow(hList, 5, 5, LOWORD(lp)-10, (HIWORD(lp) - iStatusWy) / 2 - 10, TRUE);
            MoveWindow(hEdit, 0, (HIWORD(lp) - iStatusWy) / 2, LOWORD(lp), (HIWORD(lp) - iStatusWy) / 2, TRUE);
            SendMessage(hStatus, WM_SIZE, wp, lp);
            break;
        case WM_NOTIFY:
            lpnmhdr = (LPNMHDR)lp;
            if (lpnmhdr->code == NM_CLICK) {
                lpnmlv = (LPNMLISTVIEW)lp;
                iClickNo = lpnmlv->iItem + 1;
                if (iClickNo == 0) {
                    // MessageBox(hWnd, "アイテムのないところをクリックしても無効です",
                    //    "Error", MB_OK);
                    break;
                }
                MyPopConnect(hWnd, hStatus);
                ShowContents(hEdit, hStatus, hList, iClickNo);
                MyPopDisconnect(hStatus);
            }
            break;
        case WM_COMMAND:
            switch (LOWORD(wp)) {
                case IDM_END:
                    SendMessage(hWnd, WM_CLOSE, 0, 0);
                    break;
                case IDM_SETSMTP:
                    DialogBox(hInst, "MYSETDLG", hWnd, (DLGPROC)MySettingProc);
                    break;
                case IDM_SETPOP:
                    DialogBox(hInst, "MYPOPDLG", hWnd, (DLGPROC)MyPopSetProc);
                    break;
                case IDM_SEND:
                    DialogBoxParam(hInst, "MYNEWMAIL",
                        hWnd, (DLGPROC)MyNewMailProc, (LPARAM)hStatus);
                    break;
                case IDM_RCV:
                    MyRcv(hWnd, hEdit, hStatus, hList);
                    break;
				case IDM_ABOUT:
					hIcon = LoadIcon(hInst, "MYICON");
					ShellAbout(hWnd, SZMYVERSION, SZMYSZMYCOPYRIGHT, hIcon);
					break;
            }
            break;
        case WM_CLOSE:
            id = MessageBox(hWnd, "終了してもよいですか", "終了確認", MB_YESNO | MB_ICONQUESTION);
            if (id == IDYES) {
                DestroyWindow(hStatus);
                DestroyWindow(hEdit);
                DestroyWindow(hList);
                DestroyWindow(hWnd);
            }
            break;
        case WM_DESTROY:
            PostQuitMessage(0);
            break;
        default:
            return (DefWindowProc(hWnd, msg, wp, lp));
    }
    return 0;
}

// SMTP送信関数
void MySnd(HWND hStatus, char *lpszTo, char *lpszSubject, char *lpszMail)
{
    WSADATA wsaData;
    LPHOSTENT lpHost;
    LPSERVENT lpServ;
    SOCKET s;
    int iProtocolPort;
    SOCKADDR_IN sockadd;
    char *seps = "\r", *token;

    if (WSAStartup(MAKEWORD(1, 1), &wsaData) != 0) {
        MessageBox(NULL, "エラーです", "Error", MB_OK);
        return;
    }
    if (strcmp(szServerName, "") == 0) {
        MessageBox(NULL, "サーバーアドレスを入力してください。", "サーバー", MB_OK);
        DialogBox(hInst, "MYDLG", hStatus, (DLGPROC)MyDlgProc);
        strcpy(szServerName, szStr);
    }
    lpHost = gethostbyname(szServerName);
    if (lpHost == NULL) {
        wsprintf(szStr, "%sが見つかりません", szServerName);
        MessageBox(NULL, szStr, "Error", MB_OK);
        return;
    }
    s = socket(PF_INET, SOCK_STREAM, 0);
    if (s == INVALID_SOCKET) {
        MessageBox(NULL, "ソケットをオープンできません", "Error", MB_OK);
        return;
    }
    lpServ = getservbyname("mail", NULL);
    if (lpServ == NULL) {
        MessageBox(NULL, "ポート指定がされていないので、デフォルトを使います", "OK", MB_OK);
        iProtocolPort = htons(IPPORT_SMTP);
    } else {
        iProtocolPort = lpServ->s_port;
    }
    sockadd.sin_family = AF_INET;
    sockadd.sin_port = iProtocolPort;
    sockadd.sin_addr = *((LPIN_ADDR)*lpHost->h_addr_list);
    if (connect(s, (PSOCKADDR)&sockadd, sizeof(sockadd))) {
        MessageBox(NULL, "サーバーソケットに接続失敗", "Error", MB_OK);
        return;
    }
    memset(szStrRcv, '\0', sizeof(szStrRcv));
    recv(s, szStrRcv, sizeof(szStrRcv), 0);
    SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szStrRcv);

    strcpy(szStr,  "HELO ");
    strcat(szStr, szServerName);
    strcat(szStr, "\r\n");
    send(s, szStr, strlen(szStr), 0);
    memset(szStrRcv, '\0', sizeof(szStrRcv));
    recv(s, szStrRcv, sizeof(szStrRcv), 0);
    SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szStrRcv);

    if (strcmp(szFrom, "") == 0) {
        MessageBox(NULL, "差出人メールアドレスを入力してください", "差出人", MB_OK);
        DialogBox(hInst, "MYDLG", hStatus, (DLGPROC)MyDlgProc);
        strcpy(szFrom, szStr);
    }

    wsprintf(szStr, "MAIL FROM : <%s>\r\n", szFrom);
    send(s, szStr, strlen(szStr), 0);
    memset(szStrRcv, '\0', sizeof(szStrRcv));
    recv(s, szStrRcv, sizeof(szStrRcv), 0);
    SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szStrRcv);

    wsprintf(szStr, "RCPT TO :<%s>\r\n", lpszTo);
    send(s, szStr, strlen(szStr), 0);
    memset(szStrRcv, '\0', sizeof(szStrRcv));
    recv(s, szStrRcv, sizeof(szStrRcv), 0);
    SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szStrRcv);

    strcpy(szStr, "DATA\r\n");
    send(s, szStr, strlen(szStr), 0);
    memset(szStrRcv, '\0', sizeof(szStrRcv));
    recv(s, szStrRcv, sizeof(szStrRcv), 0);
    SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szStrRcv);

    strcpy(szStr, "X-Mailer: Nekodemo_Wakaru-Mailer\r\n");
    send(s, szStr, strlen(szStr), 0);

    if (strcmp(szReplyTo, "") != 0) {
        wsprintf(szStr, "Reply-To: %s\r\n", szReplyTo);
        send(s, szStr, strlen(szStr), 0);
    }

    
    MySJisToJis(lpszSubject, szStr);
    encode(szStr, lpszSubject);
    wsprintf(szStr, "Subject: %s \r\n", lpszSubject);
    send(s, szStr, strlen(szStr), 0);
    
    //ここでもう一度「\r\n」を送信しておかないと
    //本文に日本語を使う場合１行目が文字化けします
    strcpy(szStr, "\r\n");
    send(s, szStr, strlen(szStr), 0);

    token = strtok(lpszMail, seps);
    
    while (token != NULL) {
        if (token[0] == '\n') {
            strcpy(szStr, token + 1);
        } else {
            strcpy(szStr, token);
        }
        strcat(szStr, "\r\n");
        send(s, szStr, strlen(szStr), 0);
        token = strtok(NULL, seps);
    }
    strcpy(szStr, ".\r\n");
    send(s, szStr, strlen(szStr), 0);
    memset(szStrRcv, '\0', sizeof(szStrRcv));
    recv(s, szStrRcv, sizeof(szStrRcv), 0);
    SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szStrRcv);

    strcpy(szStr,  "QUIT\r\n");
    send(s, szStr, strlen(szStr), 0);
    memset(szStrRcv, '\0', sizeof(szStrRcv));
    recv(s, szStrRcv, sizeof(szStrRcv), 0);
    SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)szStrRcv);
    closesocket(s);
    WSACleanup();
    SendMessage(hStatus, SB_SETTEXT, 0 | 0, (LPARAM)"通信を終了しました");
    return;
}

// POP3受信関数
void MyRcv(HWND hWnd, HWND hEdit, HWND hStatus, HWND hList)
{
    int iNo, iNOofMail;
    char *token, szSubject[1024], szBuf[1024],
        szFrom[128], szSize[32], szDate[32], szCopy[1024 * 50];

    Edit_SetText(hEdit, "");
    MyPopConnect(hWnd, hStatus);

    strcpy(szStr, "STAT\r\n");
    send(s, szStr, strlen(szStr), 0);
    memset(szStrRcv, '\0', sizeof(szStrRcv));
    recv(s, szStrRcv, sizeof(szStrRcv), 0);
    token = strtok(szStrRcv, " ");
    token = strtok(NULL, " ");
    iNOofMail = atoi(token);

    ListView_DeleteAllItems(hList);
    for (iNo = 1; iNo <= iNOofMail; iNo++) {
        wsprintf(szBuf, "現在第%d番目のヘッダ情報取得中", iNo);
        SendMessage(hStatus, SB_SETTEXT,(WPARAM)(0 | 0), (LPARAM)szBuf);

        //TOPコマンドを送ってヘッダを取得する
        wsprintf(szStr, "TOP %d 0\r\n", iNo);
        send(s, szStr, strlen(szStr), 0);
        memset(szStrRcv, '\0', sizeof(szStrRcv));
        recv(s, szStrRcv, sizeof(szStrRcv), 0);//１行読み飛ばし
        recv(s, szStrRcv, sizeof(szStrRcv), 0);
        memset(szSubject, '\0', sizeof(szSubject));
        memset(szBuf, '\0', sizeof(szBuf));
        memset(szFrom, '\0', sizeof(szFrom));
        memset(szDate, '\0', sizeof(szDate));
        strcpy(szCopy, szStrRcv);
        GetSubject(szCopy, szBuf);
        GetString(szBuf, szSubject);
        InsertItem(hList, iNo - 1, 0, szSubject);
        memset(szCopy, '\0', sizeof(szCopy));
        strcpy(szCopy, szStrRcv);
        GetFrom(szCopy, szFrom);
        InsertItem(hList, iNo - 1, 1, szFrom);
        strcpy(szCopy, szStrRcv);
        GetDate(szCopy, szDate);
        InsertItem(hList, iNo - 1, 2, szDate);

        wsprintf(szBuf, "LIST %d\r\n", iNo);
        send(s, szBuf, strlen(szBuf), 0);
        memset(szStrRcv, '\0', sizeof(szStrRcv));
        recv(s, szStrRcv, sizeof(szStrRcv), 0);
        wsprintf(szSize, "%d", GetMailSize(szStrRcv));
        InsertItem(hList, iNo - 1, 3, szSize);
    }
    MyPopDisconnect(hStatus);

    return;
}

void MyPopConnect(HWND hWnd, HWND hStatus)
{
    WSADATA wsaData;
    LPHOSTENT lpHost;
    SOCKADDR_IN sockadd;

    if (WSAStartup(MAKEWORD(1, 1), &wsaData) != 0) {
        MessageBox(hWnd, "エラーです", "Error", MB_OK);
        return;
    }

    while (strcmp(szPopServer, "") == 0 || 
        strcmp(szUserName, "") == 0 ||
        strcmp(szPass, "") == 0) {
        MessageBox(hWnd, "POP3の設定を行ってください。", "注意", MB_OK);
        DialogBox(hInst, "MYPOPDLG", hWnd, (DLGPROC)MyPopSetProc);
    }

    lpHost = gethostbyname(szPopServer);

    if (lpHost == NULL) {
        wsprintf(szStr, "%sが見つかりません", szPopServer);
        MessageBox(hWnd, szStr, "Error", MB_OK);
        return;
    }

    s = socket(PF_INET, SOCK_STREAM, 0);
    if (s == INVALID_SOCKET) {
        MessageBox(hWnd, "ソケットをオープンできません", "Error", MB_OK);
        return;
    }

    sockadd.sin_family = AF_INET;
    sockadd.sin_port = htons(110);	// POP3のポートは110番
    sockadd.sin_addr = *((LPIN_ADDR)*lpHost->h_addr_list);

    if (connect(s, (PSOCKADDR)&sockadd, sizeof(sockadd))) {
        MessageBox(hWnd, "サーバーソケットに接続失敗", "Error", MB_OK);
        closesocket(s);
        WSACleanup();
        return;
    }

    memset(szStrRcv, '\0', sizeof(szStrRcv));
    recv(s, szStrRcv, sizeof(szStrRcv), 0);
    SendMessage(hStatus, SB_SETTEXT,
        (WPARAM)(0 | 0), (LPARAM)szStrRcv);

    wsprintf(szStr, "USER %s\r\n", szUserName);
    send(s, szStr, strlen(szStr), 0);
    memset(szStrRcv, '\0', sizeof(szStrRcv));
    recv(s, szStrRcv, sizeof(szStrRcv), 0);
    SendMessage(hStatus, SB_SETTEXT,
        (WPARAM)(0 | 0), (LPARAM)szStrRcv);

    wsprintf(szStr, "PASS %s\r\n", szPass);
    send(s, szStr, strlen(szStr), 0);
    memset(szStrRcv, '\0', sizeof(szStrRcv));
    recv(s, szStrRcv, sizeof(szStrRcv), 0);
    SendMessage(hStatus, SB_SETTEXT,
        (WPARAM)(0 | 0), (LPARAM)szStrRcv);

    if (strstr(szStrRcv, "has 0") != NULL) {
        MessageBox(hWnd, "メッセージはありません", "No Message", MB_OK);
        return;
    }
    bPOP = TRUE;
    return;
}

// POP3に接続 
void MyPopDisconnect(HWND hStatus)
{
    strcpy(szStr, "QUIT\r\n");
    send(s, szStr, strlen(szStr), 0);
    memset(szStrRcv, '\0', sizeof(szStrRcv));
    recv(s, szStrRcv, sizeof(szStrRcv), 0);
    SendMessage(hStatus, SB_SETTEXT,
        (WPARAM)(0 | 0), (LPARAM)szStrRcv);
    closesocket(s);
    WSACleanup();
    bPOP = FALSE;
    return;
}

LRESULT CALLBACK MyNewMailProc(HWND hDlg, UINT msg, WPARAM wp, LPARAM lp)
{
	static HWND hTo, hSubject, hMail, hStatus;
	char szTo[256], szSubject[256], *lpszMail, *lpszResult;
	int iLen;
	HGLOBAL hMem, hResult;
	
	switch (msg) {
        case WM_INITDIALOG:
            hTo = GetDlgItem(hDlg, IDC_EDIT1);
            hSubject = GetDlgItem(hDlg, IDC_EDIT2);
            hMail = GetDlgItem(hDlg, IDC_EDIT3);
            hStatus = (HWND)lp;
            return TRUE;
        case WM_COMMAND:
            switch (LOWORD(wp)) {
                case IDOK:
                    Edit_GetText(hTo, szTo, sizeof(szTo));
                    Edit_GetText(hSubject, szSubject, sizeof(szSubject));
                    iLen = Edit_GetTextLength(hMail) + 1;
                    hMem = GlobalAlloc(GHND, iLen);
                    hResult = GlobalAlloc(GHND, iLen + 1024);
                    if (hMem == NULL || hResult == NULL) {
                        MessageBox(hDlg, "メモリの確保に失敗しました", "Error", MB_OK);
                        return FALSE;
                    }
                    lpszMail = (char *)GlobalLock(hMem);
                    lpszResult = (char *)GlobalLock(hResult);
                    Edit_GetText(hMail, lpszMail, iLen);
                    MySJisToJis(lpszMail, lpszResult);
                    MySnd(hStatus, szTo, szSubject, lpszResult);
                    GlobalUnlock(hMem);
                    GlobalFree(hMem);
                    GlobalUnlock(hResult);
                    GlobalFree(hResult);
                    EndDialog(hDlg, IDOK);
                    return TRUE;
                case IDCANCEL:
                    EndDialog(hDlg, IDCANCEL);
                    return TRUE;
            }
            return FALSE;
    }
    return FALSE;
}

LRESULT CALLBACK MyDlgProc(HWND hDlg, UINT msg, WPARAM wp, LPARAM lp)
{
	static HWND hEdit;
		switch (msg) {
			case WM_INITDIALOG:
				hEdit = GetDlgItem(hDlg, IDC_EDIT1);
				return TRUE;
			case WM_COMMAND:
				switch(LOWORD(wp)){
					case IDOK:
						Edit_GetText(hEdit, szStr, sizeof(szStr));
						EndDialog(hDlg, IDOK);
						return TRUE;
					case IDCANCEL:
						EndDialog(hDlg, IDCANCEL);
						return TRUE;
				}
				return FALSE;
		}
		return FALSE;
}

LRESULT CALLBACK MySettingProc(HWND hDlg, UINT msg, WPARAM wp, LPARAM lp)
{
	static HWND hServerName, hFrom, hReplyTo;

	switch(msg){
		case WM_INITDIALOG:
			hServerName = GetDlgItem(hDlg, IDC_EDIT1);
			hFrom = GetDlgItem(hDlg, IDC_EDIT2);
            hReplyTo = GetDlgItem(hDlg, IDC_EDIT2);
			Edit_SetText(hServerName, szServerName);
			Edit_SetText(hFrom, szFrom);
			Edit_SetText(hReplyTo, szReplyTo);
			return TRUE;
		case WM_COMMAND:
			switch(LOWORD(wp)){
				case IDOK:
					Edit_GetText(hServerName, szServerName, sizeof(szServerName));
					Edit_GetText(hFrom, szFrom, sizeof(szFrom));
					Edit_GetText(hReplyTo, szReplyTo, sizeof(szReplyTo));
					
					if(strcmp(szReplyTo, "") == 0)
						strcpy(szReplyTo, szFrom);
					EndDialog(hDlg, IDOK);
					return TRUE;
				case IDCANCEL:
					EndDialog(hDlg, IDCANCEL);
					return TRUE;
			}
			return FALSE;
	}
	return FALSE;
}

LRESULT CALLBACK MyPopSetProc(HWND hDlg, UINT msg, WPARAM wp, LPARAM lp)
{
	static HWND hPopServer, hUserName, hPass;
	
	switch(msg){
		case WM_INITDIALOG:
			hPopServer = GetDlgItem(hDlg, IDC_EDIT1);
			hUserName = GetDlgItem(hDlg, IDC_EDIT2);
			hPass = GetDlgItem(hDlg, IDC_EDIT3);
			Edit_SetText(hPopServer, szPopServer);
			Edit_SetText(hUserName, szUserName);
			Edit_SetText(hPass, szPass);
			return TRUE;
		case WM_COMMAND:
			switch(LOWORD(wp)){
				case IDOK:
					Edit_GetText(hPopServer, szPopServer, sizeof(szPopServer));
					Edit_GetText(hUserName, szUserName, sizeof(szUserName));
					Edit_GetText(hPass, szPass, sizeof(szPass));
					EndDialog(hDlg, IDOK);
					return TRUE;
				case IDCANCEL:
					EndDialog(hDlg, IDCANCEL);
					return TRUE;
			}
			return FALSE;
	}
	return FALSE;
}

int GetMailSize(char *szMsg)
{
	char *token;
	char seps[] = " ";

    // 半角スペースをデリミタとして３番目のトークンを取得する
    // 「+OK 番号 **** octes」という文字列から****を取得する
	token = strtok(szMsg, seps);		// 無駄読み１番目のトークン
	token = strtok(NULL, seps);			// 無駄読み２番目のトークン
	token = strtok(NULL, seps);			// ３番目のトークン
	return atoi(token);
}

void MyJisToSJis(char *lpOrg, char *lpDest)
{
	int i = 0, iR = 0, c;
	BOOL bSTART = FALSE;	// JIS->SJIS変換が必要かどうか
	
	while(1){

		if(lpOrg[i] == 0x1b && lpOrg[i + 1] == 0x24 && lpOrg[i + 2] == 0x42){
			bSTART = TRUE;
			i += 3;
			continue;
		}
		
		if(lpOrg[i] == 0x1b && lpOrg[i+ 1] == 0x28 && lpOrg[i + 2] == 0x42){
			bSTART = FALSE;
			i += 3;
			continue;
		}
		
		if(lpOrg[i] >= 0x21 && lpOrg[i] <= 0x7e && bSTART){
			c = MAKEWORD(lpOrg[i + 1], lpOrg[i]);
			lpDest[iR] = (char)HIBYTE(_mbcjistojms(c));
			lpDest[iR + 1] = (char)LOBYTE(_mbcjistojms(c));
			i += 2;
			iR += 2;
			continue;
		}
		
		if(i == 0 && lpOrg[0] == '\n'){
			lpDest[iR] = '\r';
			lpDest[iR + 1] = '\n';
			i++;
			iR += 2;
			continue;
		}
		
		if(lpOrg[i] == '\n' && lpOrg[i - 1] != '\r'){
			lpDest[iR] = '\r';
			lpDest[iR + 1] = '\n';
			i++;
			iR += 2;
			continue;
		}
		
		if(lpOrg[i] == '\0'){
			lpDest[iR] = lpOrg[i];
			break;
		}
		lpDest[iR] = lpOrg[i];
		iR++;
		i++;
	}
	return;
}

void MySJisToJis(char *lpszOrg, char *lpszDest)
{
	int i = 0, iR = 0, c;
	BOOL bSTART = FALSE;
	
	while(1){
		if(_ismbblead(lpszOrg[i]) && bSTART == FALSE){
			lpszDest[iR] = 0x1b;
			lpszDest[iR + 1] = 0x24;
			lpszDest[iR + 2] = 0x42;
			c = MAKEWORD(lpszOrg[i + 1], lpszOrg[i]);
			lpszDest[iR + 3] = (char)HIBYTE(_mbcjmstojis(c));
			lpszDest[iR + 4] = (char)LOBYTE(_mbcjmstojis(c));
			bSTART = TRUE;
			i += 2;
			iR += 5;

			if(!_ismbblead(lpszOrg[i])){
				lpszDest[iR] = 0x1b;
				lpszDest[iR + 1] = 0x28;
				lpszDest[iR + 2] = 0x42;
				bSTART = FALSE;
				iR += 3;
			}
			continue;
		}
		
		if(_ismbblead(lpszOrg[i]) && bSTART){
			c = MAKEWORD(lpszOrg[i + 1], lpszOrg[i]);
			lpszDest[iR] = (char)HIBYTE(_mbcjmstojis(c));
			lpszDest[iR + 1] = (char)LOBYTE(_mbcjmstojis(c));
			i += 2;
			iR += 2;
			if(!_ismbblead(lpszOrg[i])){
				lpszDest[iR] = 0x1b;
				lpszDest[iR + 1] = 0x28;
				lpszDest[iR + 2] = 0x42;
				bSTART = FALSE;
				iR += 3;
			}
			continue;
		}

		if(lpszOrg[i] == '\r'){
			lpszDest[iR] = '\n';
			i += 2;
			iR++;
			continue;
		}
		
		if(lpszOrg[i] == '\0'){
			lpszDest[iR] = lpszOrg[i];
			break;
		}
		lpszDest[iR] = lpszOrg[i];
		iR++;
		i++;
	}
	return;
}

void GetSubject(char *szStrRcv, char *szSubject)
{
    BOOL bCopy = FALSE;
    char *token;

    token = strtok(szStrRcv, "\r\n");
    while (token != NULL) {
        if (strstr(token, "Subject:") == token) {
            strcpy(szSubject, token + 9);
            token = strtok(NULL, "\r\n");
            while (strstr(token, ":") == NULL) {
                strcat(szSubject, token);
                token = strtok(NULL, "\r\n");
            }
        }
        token = strtok(NULL, "\r\n");
    }
    return;
}

void GetFrom(char *szStr, char *szFrom)
{
    char *token;
    char szBuf[1024];

    token = strtok(szStr, "\r\n");
    while (token != NULL) {
        if (strstr(token, "From:") != NULL) {
            strcpy(szBuf, token + 5);
            GetString(szBuf, szFrom);
        }
        token = strtok(NULL, "\r\n");
    }
    return;
}

void GetDate(char *szStrRcv, char *szDate)
{
    char *token;

    token = strtok(szStrRcv, "\r\n");
    while (token != NULL) {
        if (strstr(token, "Date:") == token) {
            strcpy(szDate, token + 5);
        }
        token = strtok(NULL, "\r\n");
    }
    return;
}

void GetString(char *szOrg, char *szResult)
{

    char *token;
    char szBuf[1024], szBuf2[1024];
    int len, mod;
    BOOL bCopy = TRUE; //?をコピーしてよいか

    token = strtok(szOrg, "=\r\n\t");
    while (token != NULL) {
        if (strstr(token, "?ISO-2022-JP?B?") != token &&
            strstr(token, "?iso-2022-jp?B?") != token) {
            if (strcmp(token, "?") != 0 && bCopy == TRUE) {
                MyJisToSJis(token, szBuf);
                strcat(szResult, szBuf);
            }
        } else {
            strcpy(szBuf, token + 15);
            if (szBuf[strlen(szBuf) - 1] == '?')
                szBuf[strlen(szBuf) - 1] = '=';
            len = strlen(szBuf);
            mod = len % 4;
            switch (mod) {
                case 1:
                    strcat(szBuf, "===");
                    bCopy = FALSE;
                    break;
                case 2:
                    strcat(szBuf, "==");
                    bCopy = FALSE;
                    break;
                case 3:
                    strcat(szBuf, "=");
                    bCopy = FALSE;
                    break;
            }
            decode(szBuf, szBuf2);
            MyJisToSJis(szBuf2, szBuf);
            strcat(szResult, szBuf);
        }
        bCopy = TRUE;
        token = strtok(NULL, "=\r\n\t");
    }
    return;
}

void decode(char *lpszStr, char *lpszResult)
{
    int len, i, iR;
    char a1, a2, a3, a4;

    len = strlen(lpszStr);
    i = 0;
    iR = 0;
    while (1) {
        if (i >= len)
            break;
        a1 = charconv(lpszStr[i]);
        a2 = charconv(lpszStr[i+1]);
        a3 = charconv(lpszStr[i+2]);
        a4 = charconv(lpszStr[i+3]);
        lpszResult[iR] = (a1 << 2) + (a2 >>4);
        lpszResult[iR + 1] = (a2 << 4) + (a3 >>2);
        lpszResult[iR + 2] = (a3 << 6) + a4;
        i += 4;
        iR += 3;
    }
    return;
}

char charconv(char c)
{
    char str[256];

    if (c >= 'A' && c <= 'Z')
        return (c - 'A');
    if (c >= 'a' && c <= 'z')
        return (c - 'a' + 0x1a);
    if (c >= '0' && c <= '9')
        return (c - '0' + 0x34);
    if (c == '+')
        return 0x3e;
    if (c == '/')
        return 0x3f;
    if (c == '=')
        return '\0';
    wsprintf(str, "不正な文字[%c]を検出しました", c);
    MessageBox(NULL, str, "Error", MB_OK);
    return '\0';
}

void encode(char *lpszOrg, char *lpszDest)
{
    int i = 0, iR = 16;

    strcpy(lpszDest, "=?ISO-2022-JP?B?");

    while (1) {
        if (lpszOrg[i] == '\0') {
            break;
        }
        lpszDest[iR] = convtobase((lpszOrg[i]) >> 2);
        if (lpszOrg[i + 1] == '\0') {
            lpszDest[iR + 1] = convtobase(((lpszOrg[i] & 0x3) << 4) );
            lpszDest[iR + 2] = '=';
            lpszDest[iR + 3] = '=';
            lpszDest[iR + 4] = '\0';
            break;
        }
        lpszDest[iR + 1] = convtobase(((lpszOrg[i] & 0x3) << 4) + ((lpszOrg[i + 1]) >> 4));
        if (lpszOrg[i + 2] == '\0') {
            lpszDest[iR + 2] = convtobase((lpszOrg[i + 1] & 0xf) << 2);
            lpszDest[iR + 3] = '=';
            lpszDest[iR + 4] = '\0';
            break;
        }
        lpszDest[iR + 2] = convtobase(((lpszOrg[i + 1] & 0xf) << 2) + ((lpszOrg[i + 2]) >> 6));
        lpszDest[iR + 3] = convtobase(lpszOrg[i + 2] & 0x3f);
        lpszDest[iR + 4] = '\0';
        i += 3;
        iR += 4;
    }
    strcat(lpszDest, "?=");
    return;
}

char convtobase(char c)
{
    if (c <= 0x19)
        return c + 'A';
    if (c >= 0x1a && c <= 0x33)
        return c - 0x1a + 'a';
    if (c >= 0x34 && c <= 0x3d)
        return c - 0x34 + '0';
    if (c == 0x3e)
        return '+';
    if (c == 0x3f)
        return '/';
    return '=';
}

void InsertColumn(HWND hList)
{
    LVCOLUMN lvcol;

    lvcol.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
    lvcol.fmt = LVCFMT_LEFT;
    lvcol.cx = 200;
    lvcol.pszText = "件名";
    lvcol.iSubItem = 0;
    ListView_InsertColumn(hList, 0, &lvcol);

    lvcol.cx = 50;
    lvcol.pszText = "差出人";
    lvcol.iSubItem = 1;
    ListView_InsertColumn(hList, 1, &lvcol);

    lvcol.cx = 100;
    lvcol.pszText = "日付";
    lvcol.iSubItem = 2;
    ListView_InsertColumn(hList, 2, &lvcol);

    lvcol.cx = 100;
    lvcol.pszText = "サイズ";
    lvcol.iSubItem = 3;
    ListView_InsertColumn(hList, 3, &lvcol);

    return;
}

void InsertItem(HWND hList, int iItemNo, int iSubNo, char *lpszText)
{
    LVITEM item;

    memset(&item, 0, sizeof(LVITEM));

    item.mask = LVIF_TEXT;
    item.pszText = lpszText;
    item.iItem = iItemNo;
    item.iSubItem = iSubNo;

    if (iSubNo == 0)
        ListView_InsertItem(hList, &item);
    else
        ListView_SetItem(hList, &item);
    return;
}

void ShowContents(HWND hEdit, HWND hStatus, HWND hList, int iClickNo)
{
    int iMailSize, iGetSize;
    char *lpszStr;
    HGLOBAL hMem;

    Edit_SetText(hEdit, "");
    wsprintf(szStr, "LIST %d\r\n", iClickNo);
    send(s, szStr, strlen(szStr), 0);
    memset(szStrRcv, '\0', sizeof(szStr));
    recv(s, szStrRcv, sizeof(szStrRcv), 0);
    SendMessage(hStatus, SB_SETTEXT,
        (WPARAM)(0 | 0), (LPARAM)szStrRcv);
            
    iMailSize = GetMailSize(szStrRcv);
    iGetSize = 0;

    hMem = GlobalAlloc(GHND, iMailSize + 1024);
    if (hMem == NULL) {
        MessageBox(NULL, "メモリを確保できません", "Error", MB_OK);
        MyPopDisconnect(hStatus);
        closesocket(s);
        WSACleanup();
        return;
    }
    lpszStr = (char *)GlobalLock(hMem);

    wsprintf(szStr, "RETR %d\r\n", iClickNo);
    send(s, szStr, strlen(szStr), 0);
    memset(szStrRcv, '\0', sizeof(szStrRcv));
    iGetSize += recv(s, szStrRcv, sizeof(szStrRcv), 0);
    while (iMailSize >= iGetSize) {
        memset(szStrRcv, '\0', sizeof(szStrRcv));
        iGetSize += recv(s, szStrRcv, sizeof(szStrRcv), 0);
        MyJisToSJis(szStrRcv, szResult);
        Edit_GetText(hEdit, lpszStr, iMailSize + 1024);
        strcat(lpszStr, szResult);
        Edit_SetText(hEdit, lpszStr);
    }
    GlobalUnlock(hMem);
    if (GlobalFree(hMem)) {
        MessageBox(NULL, "メモリの開放に失敗しました", "Error", MB_OK);
    }
    return;
}

