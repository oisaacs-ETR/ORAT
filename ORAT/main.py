import tkinter as tk
from pathlib import Path
import sys
from tkinter import ttk, scrolledtext, messagebox
import win32com.client

# ======================= SAP CONNECTION =======================


def get_all_sap_sessions():
    """Returns a list of all open SAP sessions with useful descriptions"""

    try:

        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        sessions = []

        for i in range(application.Children.Count):
            connection = application.Children(i)

            for j in range(connection.Children.Count):
                session = connection.Children(j)
                info = session.Info
                desc = f"{info.SystemName} - {info.User} @ {info.Client} ({info.SessionNumber})"
                sessions.append((session, desc))
        return sessions

    except Exception as e:
        messagebox.showerror("SAP Error", f"Cannot find SAP GUI:\n{e}")
        return []


# ======================= SAP ACTIONS =======================


def rlo_batch(session, wo_numbers, log_func):

    try:
    #resize window to ensure accuracy of scripting actions
        session.findById("wnd[0]").resizeWorkingPane(171, 39, False)
    #type IW32 into searchbar and press enter
        session.findById("wnd[0]/tbar[0]/okcd").text = "iw32"
        session.findById("wnd[0]").sendVKey(0)

        
    #LIMIT: ensure no more than the limit set by variable "limit" # of work orders are processed
        limit = 100
        if len(wo_numbers) > limit:
            raise Exception(f"Too many work orders: {len(wo_numbers)}. The limit is {limit}.")


        for wo in wo_numbers:

            log_func(f"Processing {wo}...")

            try:
        #enter WO number into IW32 and press enter
                session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = wo
                session.findById("wnd[0]").sendVKey(0)
        #click on the "User Status" field, check RLO, and uncheck RCD
                session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/btnBUTTON_STATUS").press()
                session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,6]").selected = True
                session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,7]").selected = False
                session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,6]").setFocus()
        #hotkeys for back and save
                session.findById("wnd[0]").sendVKey(3)
                session.findById("wnd[0]").sendVKey(11)

                log_func(f"✓ {wo}")

            except Exception as e:

                log_func(f"✗ {wo}: {str(e)}")

        log_func(f"\nFinished {len(wo_numbers)} work order(s).")
    #back button to return to home screen
        session.findById("wnd[0]").sendVKey(3)

    except Exception as e:

        log_func(f"RLO Error: {e}")


def complete_dd(session, disdoc, log_func):

    try:
    #resize window to ensure accuracy of scripting actions
        session.findById("wnd[0]").resizeWorkingPane(171, 39, False)
    #use searchbar to open ZAMIFLAG_UPDATE
        session.findById("wnd[0]/tbar[0]/okcd").text = "zamiflag_update"
        session.findById("wnd[0]").sendVKey(0)
    #select option to clear AMI flag and execute
        session.findById("wnd[0]/usr/radP_CLEAR").select()
        session.findById("wnd[0]/usr/ctxtS_DOC-LOW").text = disdoc
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        session.findById("wnd[1]/usr/btnBUTTON_1").press()
    #use back button twice to return to home screen
        session.findById("wnd[0]").sendVKey(3)
        session.findById("wnd[0]").sendVKey(3)
        log_func("AMI flag cleared")

    except:
        pass

    try:
    #use searchbar to open EC86
        session.findById("wnd[0]/tbar[0]/okcd").text = "ec86"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtEDISCD-DISCNO").text = disdoc
    #Key shortcuts to enter disconnection, enter reconnetion, then save twice
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(6)
        session.findById("wnd[0]").sendVKey(11)
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]").sendVKey(11)
        session.findById("wnd[0]").sendVKey(11)
    #back button to return to home screen
        session.findById("wnd[0]").sendVKey(3)
        log_func(f"DD# {disdoc} completed")
        log_func(f"Please ensure that all work orders associated with DD# {disdoc} are technically completed")

    except Exception as e:

        log_func(f"EC86 error: {e}")


def complete_wo(session, wo, log_func):

    try:
    #resize window to ensure accuracy of scripting actions
        session.findById("wnd[0]").resizeWorkingPane(171, 39, False)
        session.findById("wnd[0]/tbar[0]/okcd").text = "iw32"
        session.findById("wnd[0]").sendVKey(0)
    #enter WO number into IW32 and press enter
        session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = wo
        session.findById("wnd[0]").sendVKey(0)
    #select the "User Status" box, then check CP, RLO, AUCP, and uncheck P and RCD
        session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/subSUB_KOPF:SAPLCOIH:1102/btnBUTTON_STATUS").press()
        session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,8]").selected = True
        session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,6]").selected = True
        session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,7]").selected = False
        session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,7]").setFocus()
        session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO").verticalScrollbar.position = 1
        session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO").verticalScrollbar.position = 2
        session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO").verticalScrollbar.position = 3
        session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO").verticalScrollbar.position = 4
        session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO").verticalScrollbar.position = 5
        session.findById("wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,10]").selected = True
        session.findById("wnd[0]").sendVKey(3)
    #TECO, save, and back to return to home screen
        session.findById("wnd[0]/tbar[1]/btn[36]").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[0]").sendVKey(3)
        log_func(f"Work Order # {wo} completed")

    except Exception as e:

        log_func(f"Error completing WO: {e}")


# ======================= GUI =======================


class ORAT(tk.Tk):

    
    def __init__(self):
        super().__init__()
        self.title("ORAT v0.5")
        self.geometry("880x740")
        self.configure(bg="#1e1e1e")
        base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
        self.iconbitmap(str(base_dir / "orat_logo.ico"))
        self.session = None
        self.all_sessions = []
        self.create_widgets()


    def log(self, msg):

        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)

    def refresh_sessions(self):

        self.all_sessions = get_all_sap_sessions()
        self.session_combo['values'] = [desc for _, desc in self.all_sessions]

        if self.all_sessions:
            self.session_combo.current(0)
            self.status_label.config(text="Ready - Select session",
                                     fg="#00ff00")

        else:
            self.session_combo['values'] = ["No SAP sessions found"]
            self.status_label.config(text="No SAP open", fg="red")

    def select_session(self):

        idx = self.session_combo.current()

        if idx >= 0 and idx < len(self.all_sessions):
            self.session = self.all_sessions[idx][0]
            self.status_label.config(
                text=f"Connected: {self.all_sessions[idx][1]}", fg="#00ff00")
            self.log("Session selected and ready!")

        else:
            self.session = None
            self.status_label.config(text="Invalid selection", fg="red")

    def create_widgets(self):

        pad = {"padx": 10, "pady": 8}

        # ----- Session Selector -----

        sess_frame = tk.LabelFrame(self,
                                   text=" Select SAP Session ",
                                   bg="#1e1e1e",
                                   fg="#00ff00")

        sess_frame.pack(fill="x", **pad)

        self.session_combo = ttk.Combobox(sess_frame,
                                          state="readonly",
                                          width=50)

        self.session_combo.pack(side="left", padx=10, pady=5)

        ttk.Button(sess_frame,
                   text="Refresh Sessions",
                   command=self.refresh_sessions).pack(side="left", padx=5)

        ttk.Button(sess_frame,
                   text="Connect to Session",
                   command=self.select_session).pack(side="left", padx=5)

        self.status_label = tk.Label(sess_frame,
                                     text="Click Refresh",
                                     bg="#1e1e1e",
                                     fg="orange",
                                     font=("Arial", 10, "bold"))

        self.status_label.pack(side="left", padx=20)

        # ----- RLO Batch -----

        rlo_frame = tk.LabelFrame(self,
                                  text=" RLO - Batch Release Work Orders ",
                                  bg="#1e1e1e",
                                  fg="#00ff00")

        rlo_frame.pack(fill="both", expand=True, **pad)

        self.rlo_text = scrolledtext.ScrolledText(rlo_frame,
                                                  height=12,
                                                  bg="#2d2d2d",
                                                  fg="white",
                                                  insertbackground="white")

        self.rlo_text.pack(fill="both", expand=True, padx=10, pady=5)

        self.rlo_text.insert(tk.END,
                             "Paste work orders here (one per line)...")

        ttk.Button(rlo_frame, text="Run RLO Batch",
                   command=self.run_rlo).pack(pady=8)

        # ----- Complete DD -----

        dd_frame = tk.Frame(self, bg="#1e1e1e")
        dd_frame.pack(fill="x", **pad)

        tk.Label(dd_frame, text="DD#:", bg="#1e1e1e",
                 fg="white").pack(side="left", padx=10)

        self.dd_entry = tk.Entry(dd_frame,
                                 width=20,
                                 bg="#333333",
                                 fg="white",
                                 insertbackground="white")

        self.dd_entry.pack(side="left", padx=5)

        ttk.Button(dd_frame, text="Complete DD",
                   command=self.run_dd).pack(side="left", padx=10)

        # ----- CPWO -----

        cpwo_frame = tk.Frame(self, bg="#1e1e1e")
        cpwo_frame.pack(fill="x", **pad)

        tk.Label(cpwo_frame, text="WO#:", bg="#1e1e1e",
                 fg="white").pack(side="left", padx=10)

        self.cpwo_entry = tk.Entry(cpwo_frame,
                                   width=20,
                                   bg="#333333",
                                   fg="white",
                                   insertbackground="white")

        self.cpwo_entry.pack(side="left", padx=5)

        ttk.Button(cpwo_frame, text="TECO Work Order",
                   command=self.run_cpwo).pack(side="left", padx=10)

        # ----- Log -----

        log_frame = tk.LabelFrame(self,
                                  text=" Log Output ",
                                  bg="#1e1e1e",
                                  fg="#00ff00")

        log_frame.pack(fill="both", expand=True, **pad)

        self.log_text = scrolledtext.ScrolledText(log_frame,
                                                  height=12,
                                                  bg="#111111",
                                                  fg="#00ff00",
                                                  insertbackground="white")

        self.log_text.pack(fill="both", expand=True, padx=10, pady=5)

        # Auto-refresh on start

        self.after(1000, self.refresh_sessions)

    def run_rlo(self):

        if not self.session:
            messagebox.showwarning("No Session", "Select a SAP session first")
            return

        text = self.rlo_text.get("1.0", tk.END).strip()

        if not text or "paste" in text.lower():
            messagebox.showwarning("Empty", "Paste work orders first")
            return

        wo_list = [line.strip() for line in text.splitlines() if line.strip()]

        self.log(f"Starting RLO batch - {len(wo_list)} order(s)")

        rlo_batch(self.session, wo_list, self.log)

    def run_dd(self):

        if not self.session:
            messagebox.showwarning("No Session", "Select a SAP session first")
            return

        dd = self.dd_entry.get().strip()

        if not dd:
            messagebox.showwarning("Empty", "Enter DD#")
            return

        self.log(f"Completing DD {dd}...")

        complete_dd(self.session, dd, self.log)

    def run_cpwo(self):

        if not self.session:
            messagebox.showwarning("No Session", "Select a SAP session first")
            return

        wo = self.cpwo_entry.get().strip()

        if not wo:
            messagebox.showwarning("Empty", "Enter WO#")
            return

        self.log(f"Running CPWO on {wo}...")

        complete_wo(self.session, wo, self.log)


if __name__ == "__main__":
    app = ORAT()
    app.mainloop()
