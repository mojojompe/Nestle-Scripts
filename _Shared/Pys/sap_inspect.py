"""
sap_inspect.py  --  read-only snapshot of every live SAP session
================================================================
Attaches to the ALREADY-RUNNING SAP GUI and prints, per connection/session:
  * window title (wnd[0].Text)
  * whether the Easy Access command field (okcd) is present
  * whether it's on the LOGON screen (RSYST-BNAME present)
  * any modal popup (wnd[1]) and its title
  * SystemName / User / Transaction / Busy

Run it WHILE SAP is open (the same state where the engine failed) so we can see
WHY 'command field not ready' fires.

    python sap_inspect.py

READ-ONLY: opens/closes/changes nothing.
"""
import sys
try:
    import win32com.client
except Exception as e:
    print("pywin32 required:", e); sys.exit(1)


def prop(o, name, default="<n/a>"):
    try:
        return getattr(o, name)
    except Exception:
        return default


def exists(session, path):
    try:
        session.findById(path)
        return True
    except Exception:
        return False


def main() -> int:
    try:
        app = win32com.client.GetObject("SAPGUI").GetScriptingEngine
    except Exception as e:
        print("Could not attach to SAP scripting engine:", e)
        print("Make sure SAP is OPEN and scripting is enabled.")
        return 1

    try:
        ncon = int(app.Children.Count)
    except Exception as e:
        print("Engine attached but no connections readable:", e)
        return 1

    print("=" * 70)
    print(f"LIVE SAP SNAPSHOT  |  connections open: {ncon}")
    print("=" * 70)

    for i in range(ncon):
        try:
            conn = app.Children(i)
        except Exception as e:
            print(f"\n[connection {i}] unreadable: {e}")
            continue
        print(f"\n==== CONNECTION [{i}] ====")
        print("  Id          :", prop(conn, "Id"))
        print("  Description :", prop(conn, "Description"))
        print("  ConnString  :", prop(conn, "ConnectionString"))
        try:
            ns = int(conn.Children.Count)
        except Exception:
            ns = 0
        print("  Sessions    :", ns)

        for j in range(ns):
            try:
                s = conn.Children(j)
            except Exception:
                print(f"    session[{j}] unreadable"); continue
            info = prop(s, "Info", None)
            sysname = prop(info, "SystemName", "") if info else ""
            user = prop(info, "User", "") if info else ""
            tcode = prop(info, "Transaction", "") if info else ""
            title = "<n/a>"
            try:
                title = s.findById("wnd[0]").Text
            except Exception:
                pass
            has_okcd = exists(s, "wnd[0]/tbar[0]/okcd")
            on_logon = exists(s, "wnd[0]/usr/txtRSYST-BNAME")
            has_popup = exists(s, "wnd[1]")
            popup_title = ""
            if has_popup:
                try:
                    popup_title = s.findById("wnd[1]").Text
                except Exception:
                    popup_title = "<?>"
            print(f"    ----- session[{j}] -----")
            print(f"      title     : {title}")
            print(f"      okcd(cmd) : {'YES' if has_okcd else 'NO'}")
            print(f"      logon-scr : {'YES' if on_logon else 'no'}")
            print(f"      popup     : {'YES -> ' + popup_title if has_popup else 'no'}")
            print(f"      sys/user  : {sysname} / {user}")
            print(f"      tcode     : {tcode}   busy={prop(s,'Busy')}")

    print("\n" + "=" * 70)
    print("READ: if okcd=NO and popup=YES, a dialog is blocking the command "
          "field. If logon-scr=YES, SSO didn't complete on that session.")
    print("=" * 70)
    return 0


if __name__ == "__main__":
    sys.exit(main())
