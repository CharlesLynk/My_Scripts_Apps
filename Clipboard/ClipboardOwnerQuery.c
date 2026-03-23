
// clipboard_owner_loop.c
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <stdio.h>

//
// We’ll prefer QueryFullProcessImageName (Kernel32) to get the image path.
// It’s available on Vista+; for completeness we fall back to GetModuleFileNameEx (PSAPI)
// if QueryFullProcessImageName fails.
//
// MSVC users: link with user32.lib, psapi.lib, kernel32.lib (kernel32 is automatic)
// MinGW users: add -luser32 -lpsapi
//

// Helper: get process image path into a UTF-8 buffer for display
static int get_process_image_path(DWORD pid, char *out, DWORD outSize) {
    if (!out || outSize == 0) return 0;

    int success = 0;
    HANDLE hProc = OpenProcess(
        PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_VM_READ,
        FALSE, pid);

    if (!hProc) return 0;

    // Try QueryFullProcessImageNameA first
    typedef BOOL (WINAPI *QFPIA)(HANDLE, DWORD, LPSTR, PDWORD);
    HMODULE hKernel = GetModuleHandleA("kernel32.dll");
    QFPIA qfpia = (QFPIA)GetProcAddress(hKernel, "QueryFullProcessImageNameA");

    if (qfpia) {
        DWORD len = outSize;
        if (qfpia(hProc, 0, out, &len)) {
            success = 1;
        }
    }

    // Fallback to PSAPI GetModuleFileNameExA if needed
    if (!success) {
        HMODULE hPsapi = LoadLibraryA("psapi.dll");
        if (hPsapi) {
            typedef DWORD (WINAPI *GMFNEA)(HANDLE, HMODULE, LPSTR, DWORD);
            GMFNEA gmfnea = (GMFNEA)GetProcAddress(hPsapi, "GetModuleFileNameExA");
            if (gmfnea) {
                if (gmfnea(hProc, NULL, out, outSize)) {
                    success = 1;
                }
            }
            FreeLibrary(hPsapi);
        }
    }

    CloseHandle(hProc);
    return success;
}

// Query once: prints PID, window title, and image path
static void query_clipboard_owner(void) {
    HWND hwnd = GetClipboardOwner();
    if (hwnd == NULL) {
        puts("No clipboard owner or clipboard not available.");
        return;
    }

    // PID
    DWORD pid = 0;
    GetWindowThreadProcessId(hwnd, &pid);

    // Window title (ANSI; fine for most cases; use wide APIs if you need full Unicode)
    char title[512] = {0};
    int tlen = GetWindowTextA(hwnd, title, (int)sizeof(title));
    if (tlen <= 0) {
        // Some windows are ownerless for text or are hidden; not fatal
        strcpy_s(title, sizeof(title), "<no title>");
    }

    // Executable path
    char img[1024] = {0};
    if (!get_process_image_path(pid, img, (DWORD)sizeof(img))) {
        strcpy_s(img, sizeof(img), "<image path unavailable>");
    }

    printf("PID: %lu\n", (unsigned long)pid);
    printf("Window title: %s\n", title);
    printf("Process image: %s\n", img);
}

int main(void) {
    puts("Clipboard owner inspector");
    puts("Press Enter to re-query. Type 'q' then Enter to quit.\n");

    // Initial query
    query_clipboard_owner();

    // Loop: Enter -> refresh, 'q' -> quit
    char line[64];
    for (;;) {
        fputs("\n[Enter=refresh | q=quit] > ", stdout);
        if (!fgets(line, sizeof line, stdin)) {
            break; // EOF or input error
        }
        if (line[0] == 'q' || line[0] == 'Q') {
            break;
        }
        query_clipboard_owner();
    }

    puts("Exiting.");
    return 0;
}
