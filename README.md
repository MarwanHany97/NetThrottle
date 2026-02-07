# NetThrottle

A simple network throttler for Windows. Think NetLimiter but just the parts you actually use — limit download/upload speed per process or block traffic entirely.

![Windows](https://img.shields.io/badge/platform-Windows%2010%2F11-blue)
![.NET 8](https://img.shields.io/badge/.NET-8.0-purple)

## What it does

- Throttle download and/or upload speed for any process (in KB/s)
- Block all network traffic for a process with one checkbox  
- Global master limiter that caps total bandwidth across everything
- Processes with the same name (like Chrome's 47 sub-processes) get grouped together — control them all at once or expand to tweak individually
- Live speed display (5-second rolling average) so you can see whats actually happening
- Minimizes to system tray when the interceptor is running so it stays out of your way

## How it works

Under the hood it uses [WinDivert](https://reqrypt.org/windivert.html) to intercept packets at the network layer. For each packet it looks up which process owns the connection using the Windows IP Helper API, checks if there are any rules for that process, and either lets the packet through, drops it (for blocking), or uses a token bucket algorithm to enforce rate limits.

When you throttle a TCP connection by dropping packets, TCP congestion control kicks in and naturally slows down. Its not byte-accurate but it works surprisingly well in practice.

## Building

You need:
- [.NET 8 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/8.0) (the SDK, not just the runtime)
- [WinDivert 2.2](https://reqrypt.org/windivert.html)

**Run as administrator** — WinDivert needs admin to install its kernel driver.

## Usage

1. Open publish -> open NetThrottle.exe
2. Hit **Start** to begin intercepting
3. Check the boxes and type in speeds for whatever you want to limit
4. Click the **▶** arrow on grouped processes to expand individual PIDs
5. The **GLOBAL** row at the top caps total bandwidth for everything
6. Close the window while the interceptor is running and it goes to the system tray — double click to reopen, right click > Exit to fully quit

### Adaptive Mode
Check the **Adaptive** box on any row and the app will dynamically adjust the actual throttle rate to make the 5-second rolling average match your target. Without it, the token bucket allows bursts that can push the average above your limit. With it, the app continuously tightens/loosens the rate until the average converges. Takes a few seconds to settle.

## Notes

- Requires Windows 10/11 x64
- Needs admin privileges (WinDivert installs a kernel driver)
- Some antivirus software flags WinDivert — its a false positive, you may need to add an exception
- Only handles IPv4 right now
- Closing the app while the interceptor is running sends it to the tray. Closing while stopped exits cleanly and removes all rules

## Credits

Built with [WinDivert](https://reqrypt.org/windivert.html) by basil00 — does all the heavy lifting of actually capturing and reinjecting packets.

AI was used extensively in the making of this program (So if you have technical questions keep that in mind i am not an expert i just needed a FOSS alternative)

## License

MIT
