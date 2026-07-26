
import os
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
from graph_builder import build_graph, repack
from llm_agent import propose_edit
from edit_engine import apply_edit, EditError
from propagate import propagate

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class DocEditGUI:
    def __init__(self):
        self.root=ctk.CTk()
        self.root.title("AI Contract Editor")
        self.root.geometry("900x650")
        self.doc_path=""
        ctk.CTkLabel(self.root,text="AI Contract Editor",font=("Arial",28,"bold")).pack(pady=20)
        self.path_label=ctk.CTkLabel(self.root,text="No document selected",wraplength=700)
        self.path_label.pack()
        ctk.CTkButton(self.root,text="Browse DOCX",command=self.browse).pack(pady=10)
        ctk.CTkLabel(self.root,text="Instruction").pack()
        self.prompt_box=ctk.CTkTextbox(self.root,width=760,height=120)
        self.prompt_box.pack(pady=10)
        ctk.CTkButton(self.root,text="Run Edit",command=self.start_pipeline).pack(pady=10)
        ctk.CTkLabel(self.root,text="Logs").pack()
        self.logs=ctk.CTkTextbox(self.root,width=760,height=240)
        self.logs.pack(pady=10)
    def log(self,t):
        self.logs.insert("end",t+"\n")
        self.logs.see("end")
    def browse(self):
        p=filedialog.askopenfilename(filetypes=[("Word Documents","*.docx")])
        if p:
            self.doc_path=p
            self.path_label.configure(text=p)
    def start_pipeline(self):
        threading.Thread(target=self.pipeline,daemon=True).start()
    def pipeline(self):
        try:
            if not self.doc_path:
                messagebox.showerror("Error","Please choose a DOCX file."); return
            ins=self.prompt_box.get("1.0","end").strip()
            if not ins:
                messagebox.showerror("Error","Please enter an instruction."); return
            self.log("Building graph...")
            g=build_graph(self.doc_path)
            op=propose_edit(g,ins)
            self.log(str(op))
            if op.get("op")=="clarify":
                messagebox.showinfo("Clarification",op.get("question","Need more information")); return
            apply_edit(g,op)
            changed={op["bookmark"]} if "bookmark" in op else set()
            for l in propagate(g,changed):
                self.log(l)
            out=os.path.join(os.path.dirname(self.doc_path),"result_contract.docx")
            repack("unpacked",out)
            self.log("Saved: "+out)
            messagebox.showinfo("Success","Saved to:\n"+out)
        except EditError as e:
            messagebox.showerror("Edit Error",str(e))
        except Exception as e:
            messagebox.showerror("Error",str(e))

if __name__=="__main__":
    app=DocEditGUI()
    app.root.mainloop()
