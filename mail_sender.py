# -*- coding: utf-8 -*-
"""
邮件群发工具（批量个性化版）
- 支持 Gmail / Outlook / QQ / 163 / 自定义 SMTP
- 名单支持 .xlsx / .csv，第一行为表头
- 正文和主题里可以用 {列名} 占位符，按名单逐行替换
  例如正文写：亲爱的 {姓名} 同学，……
"""
import csv
import json
import os
import queue
import re
import smtplib
import ssl
import threading
import time
import tkinter as tk
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from tkinter import filedialog, messagebox, scrolledtext, ttk

try:
    import openpyxl
except ImportError:
    openpyxl = None

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "config.json")

SMTP_PRESETS = {
    "Gmail（smtp.gmail.com:465 SSL）": ("smtp.gmail.com", 465, "SSL"),
    "Outlook / Microsoft 365（smtp.office365.com:587 STARTTLS）": ("smtp.office365.com", 587, "STARTTLS"),
    "QQ 邮箱（smtp.qq.com:465 SSL）": ("smtp.qq.com", 465, "SSL"),
    "163 邮箱（smtp.163.com:465 SSL）": ("smtp.163.com", 465, "SSL"),
    "自定义": ("", 465, "SSL"),
}

PLACEHOLDER_RE = re.compile(r"\{([^{}]+)\}")


def render_template(text, row):
    """把 {列名} 替换成该行对应的值，未知的占位符原样保留。"""
    def sub(m):
        key = m.group(1).strip()
        if key in row:
            return str(row[key])
        return m.group(0)
    return PLACEHOLDER_RE.sub(sub, text)


def load_recipients(path):
    """读取名单文件，返回 (表头列表, 行列表[dict])。自动跳过空邮箱行。"""
    ext = os.path.splitext(path)[1].lower()
    rows = []
    if ext in (".xlsx", ".xlsm"):
        if openpyxl is None:
            raise RuntimeError("读取 Excel 需要 openpyxl 库，请改用 CSV 或重新安装环境")
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        it = ws.iter_rows(values_only=True)
        header = None
        for r in it:
            cells = ["" if c is None else str(c).strip() for c in r]
            if any(cells):
                if header is None:
                    header = cells
                else:
                    rows.append(dict(zip(header, cells)))
        wb.close()
    else:
        # CSV：自动尝试 utf-8-sig / gbk
        for enc in ("utf-8-sig", "gbk", "utf-8"):
            try:
                with open(path, "r", encoding=enc, newline="") as f:
                    reader = csv.reader(f)
                    data = [r for r in reader if any(c.strip() for c in r)]
                break
            except (UnicodeDecodeError, UnicodeError):
                data = None
        if not data:
            raise RuntimeError("CSV 编码无法识别，请另存为「CSV UTF-8」格式")
        header = [h.strip() for h in data[0]]
        for r in data[1:]:
            r = [c.strip() for c in r] + [""] * (len(header) - len(r))
            rows.append(dict(zip(header, r)))
    # 过滤没有邮箱的行
    if header and "邮箱" in header:
        rows = [r for r in rows if r.get("邮箱", "").strip()]
    return header, rows


class MailSenderApp:
    def __init__(self, root):
        self.root = root
        root.title("邮件群发工具 · 批量个性化")
        root.geometry("860x780")
        root.minsize(760, 640)

        self.log_q = queue.Queue()
        self.stop_flag = threading.Event()
        self.sending = False
        self.header = []
        self.rows = []

        self._build_ui()
        self._load_config()
        self.root.after(100, self._poll_log)

    # ---------- UI ----------
    def _build_ui(self):
        pad = dict(padx=6, pady=3)

        frm_smtp = ttk.LabelFrame(self.root, text="① 发件邮箱设置（SMTP）")
        frm_smtp.pack(fill="x", **pad)

        ttk.Label(frm_smtp, text="服务商预设:").grid(row=0, column=0, sticky="e", **pad)
        self.cmb_preset = ttk.Combobox(frm_smtp, values=list(SMTP_PRESETS), width=52, state="readonly")
        self.cmb_preset.current(0)
        self.cmb_preset.grid(row=0, column=1, columnspan=3, sticky="w", **pad)
        self.cmb_preset.bind("<<ComboboxSelected>>", self._on_preset)

        ttk.Label(frm_smtp, text="SMTP 服务器:").grid(row=1, column=0, sticky="e", **pad)
        self.ent_host = ttk.Entry(frm_smtp, width=34)
        self.ent_host.insert(0, SMTP_PRESETS[self.cmb_preset.get()][0])
        self.ent_host.grid(row=1, column=1, sticky="w", **pad)

        ttk.Label(frm_smtp, text="端口:").grid(row=1, column=2, sticky="e", **pad)
        self.ent_port = ttk.Entry(frm_smtp, width=8)
        self.ent_port.insert(0, str(SMTP_PRESETS[self.cmb_preset.get()][1]))
        self.ent_port.grid(row=1, column=3, sticky="w", **pad)

        ttk.Label(frm_smtp, text="加密方式:").grid(row=2, column=0, sticky="e", **pad)
        self.cmb_sec = ttk.Combobox(frm_smtp, values=["SSL", "STARTTLS", "无"], width=10, state="readonly")
        self.cmb_sec.set(SMTP_PRESETS[self.cmb_preset.get()][2])
        self.cmb_sec.grid(row=2, column=1, sticky="w", **pad)

        ttk.Label(frm_smtp, text="发件邮箱:").grid(row=3, column=0, sticky="e", **pad)
        self.ent_user = ttk.Entry(frm_smtp, width=34)
        self.ent_user.grid(row=3, column=1, sticky="w", **pad)

        ttk.Label(frm_smtp, text="授权码/密码:").grid(row=4, column=0, sticky="e", **pad)
        self.ent_pass = ttk.Entry(frm_smtp, width=34, show="*")
        self.ent_pass.grid(row=4, column=1, sticky="w", **pad)

        ttk.Label(frm_smtp, text="发件人显示名:").grid(row=5, column=0, sticky="e", **pad)
        self.ent_name = ttk.Entry(frm_smtp, width=34)
        self.ent_name.grid(row=5, column=1, sticky="w", **pad)

        self.var_remember = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm_smtp, text="记住授权码（明文保存在本机 config.json）",
                        variable=self.var_remember).grid(row=5, column=2, columnspan=2, sticky="w", **pad)

        frm_list = ttk.LabelFrame(self.root, text="② 收件人名单（Excel/CSV，第一行表头，需含「邮箱」列）")
        frm_list.pack(fill="x", **pad)
        self.ent_list = ttk.Entry(frm_list, width=72)
        self.ent_list.pack(side="left", padx=6, pady=6, fill="x", expand=True)
        ttk.Button(frm_list, text="浏览…", command=self._browse_list).pack(side="left", padx=4)
        ttk.Button(frm_list, text="读取名单", command=self._load_list).pack(side="left", padx=4)
        ttk.Button(frm_list, text="生成模板", command=self._make_template).pack(side="left", padx=4)
        self.lbl_list_info = ttk.Label(frm_list, text="未读取名单")
        self.lbl_list_info.pack(side="left", padx=8)

        frm_mail = ttk.LabelFrame(self.root, text="③ 邮件内容（可用 {列名} 占位符，如 {姓名}）")
        frm_mail.pack(fill="both", expand=True, **pad)
        ttk.Label(frm_mail, text="主题:").grid(row=0, column=0, sticky="e", **pad)
        self.ent_subject = ttk.Entry(frm_mail, width=80)
        self.ent_subject.grid(row=0, column=1, columnspan=2, sticky="we", **pad)

        self.var_html = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm_mail, text="正文是 HTML 格式（支持加粗、颜色、图片链接等）",
                        variable=self.var_html).grid(row=1, column=1, sticky="w", **pad)

        self.txt_body = scrolledtext.ScrolledText(frm_mail, height=10, font=("Microsoft YaHei UI", 10))
        self.txt_body.grid(row=2, column=0, columnspan=3, sticky="nsew", **pad)
        self.txt_body.insert("1.0", "亲爱的 {姓名}：\n\n    ……（正文第一行会自动缩进，直接写内容即可）\n\n祝好！")
        frm_mail.columnconfigure(1, weight=1)
        frm_mail.rowconfigure(2, weight=1)

        frm_send = ttk.LabelFrame(self.root, text="④ 发送")
        frm_send.pack(fill="x", **pad)
        ttk.Label(frm_send, text="每封间隔(秒):").pack(side="left", padx=6)
        self.spn_delay = ttk.Spinbox(frm_send, from_=0, to=120, width=5)
        self.spn_delay.set(3)
        self.spn_delay.pack(side="left")
        ttk.Button(frm_send, text="发送测试邮件（发给自己）", command=self._send_test).pack(side="left", padx=10)
        self.btn_send = ttk.Button(frm_send, text="开始群发", command=self._start_send)
        self.btn_send.pack(side="left", padx=6)
        self.btn_stop = ttk.Button(frm_send, text="停止", command=self.stop_flag.set, state="disabled")
        self.btn_stop.pack(side="left", padx=6)
        self.lbl_progress = ttk.Label(frm_send, text="待命")
        self.lbl_progress.pack(side="left", padx=12)

        frm_log = ttk.LabelFrame(self.root, text="发送日志")
        frm_log.pack(fill="both", expand=True, **pad)
        self.txt_log = scrolledtext.ScrolledText(frm_log, height=9, font=("Consolas", 9))
        self.txt_log.pack(fill="both", expand=True, padx=4, pady=4)

    # ---------- 配置 ----------
    def _load_config(self):
        if not os.path.exists(CONFIG_PATH):
            return
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            self.ent_host.delete(0, "end"); self.ent_host.insert(0, cfg.get("host", ""))
            self.ent_port.delete(0, "end"); self.ent_port.insert(0, str(cfg.get("port", 465)))
            self.cmb_sec.set(cfg.get("sec", "SSL"))
            self.ent_user.insert(0, cfg.get("user", ""))
            self.ent_name.insert(0, cfg.get("name", ""))
            self.ent_subject.insert(0, cfg.get("subject", ""))
            if cfg.get("remember_pass") and cfg.get("pass"):
                self.ent_pass.insert(0, cfg["pass"])
                self.var_remember.set(True)
            if cfg.get("list_path"):
                self.ent_list.insert(0, cfg["list_path"])
            if cfg.get("body"):
                self.txt_body.delete("1.0", "end")
                self.txt_body.insert("1.0", cfg["body"])
        except Exception:
            pass

    def _save_config(self):
        cfg = {
            "host": self.ent_host.get().strip(),
            "port": self.ent_port.get().strip(),
            "sec": self.cmb_sec.get(),
            "user": self.ent_user.get().strip(),
            "name": self.ent_name.get().strip(),
            "subject": self.ent_subject.get().strip(),
            "body": self.txt_body.get("1.0", "end").rstrip("\n"),
            "list_path": self.ent_list.get().strip(),
            "remember_pass": self.var_remember.get(),
            "pass": self.ent_pass.get() if self.var_remember.get() else "",
        }
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ---------- 事件 ----------
    def _on_preset(self, _=None):
        host, port, sec = SMTP_PRESETS[self.cmb_preset.get()]
        self.ent_host.delete(0, "end"); self.ent_host.insert(0, host)
        self.ent_port.delete(0, "end"); self.ent_port.insert(0, str(port))
        self.cmb_sec.set(sec)

    def _browse_list(self):
        path = filedialog.askopenfilename(
            title="选择名单文件",
            filetypes=[("表格文件", "*.xlsx *.csv"), ("所有文件", "*.*")])
        if path:
            self.ent_list.delete(0, "end")
            self.ent_list.insert(0, path)
            self._load_list()

    def _load_list(self):
        path = self.ent_list.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("提示", "请先选择有效的名单文件")
            return
        try:
            self.header, self.rows = load_recipients(path)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return
        if not self.header or "邮箱" not in self.header:
            messagebox.showerror("名单格式不对",
                                 "名单第一行表头里必须有一列叫「邮箱」。\n"
                                 "当前表头：%s" % (", ".join(self.header) or "空"))
            self.rows = []
            return
        self.lbl_list_info.config(text="已读取 %d 位收件人，可用占位符：{%s}" %
                                  (len(self.rows), "} {".join(self.header)))

    def _make_template(self):
        path = os.path.join(APP_DIR, "名单模板.xlsx")
        if openpyxl is None:
            path = os.path.join(APP_DIR, "名单模板.csv")
            with open(path, "w", encoding="utf-8-sig", newline="") as f:
                f.write("邮箱,姓名,班级\n")
            self._log("已生成模板：%s（无 openpyxl，仅 CSV）" % path)
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "名单"
        ws.append(["邮箱", "姓名", "班级"])
        ws.append(["someone@example.com", "张三", "计算机2301"])
        ws.append(["another@example.com", "李四", "电子2302"])
        for col, w in zip("ABC", (28, 12, 16)):
            ws.column_dimensions[col].width = w
        wb.save(path)
        self._log("已生成模板：%s" % path)
        messagebox.showinfo("模板已生成", "已生成「名单模板.xlsx」\n第一行是表头（必须保留「邮箱」列），下面每行一个收件人。\n除邮箱外可以随意增删列，列名都能用作占位符。")

    # ---------- 发送 ----------
    def _collect_settings(self):
        host = self.ent_host.get().strip()
        port = int(self.ent_port.get().strip() or "465")
        sec = self.cmb_sec.get()
        user = self.ent_user.get().strip()
        pwd = self.ent_pass.get()
        disp = self.ent_name.get().strip() or user
        subject = self.ent_subject.get().strip()
        body = self.txt_body.get("1.0", "end").rstrip("\n")
        if not host:
            raise ValueError("请填写 SMTP 服务器地址")
        if not user:
            raise ValueError("请填写发件邮箱")
        if not pwd:
            raise ValueError("请填写授权码/密码（不是邮箱登录密码）")
        if not subject:
            raise ValueError("请填写邮件主题")
        if not body.strip():
            raise ValueError("请填写邮件正文")
        return host, port, sec, user, pwd, disp, subject, body

    def _connect(self, host, port, sec):
        if sec == "SSL":
            server = smtplib.SMTP_SSL(host, port, timeout=30,
                                      context=ssl.create_default_context())
        else:
            server = smtplib.SMTP(host, port, timeout=30)
            if sec == "STARTTLS":
                server.starttls(context=ssl.create_default_context())
        server.login(self.ent_user.get().strip(), self.ent_pass.get())
        return server

    def _build_message(self, user, disp, subject, body, to_addr, row, as_html):
        msg = MIMEMultipart("alternative")
        msg["From"] = formataddr((str(Header(disp, "utf-8")), user))
        msg["To"] = to_addr
        msg["Subject"] = Header(render_template(subject, row), "utf-8")
        if as_html:
            html = "<html><body><div style='white-space:pre-wrap;font-family:Microsoft YaHei,Segoe UI,sans-serif;font-size:14px;'>%s</div></body></html>" % body
            msg.attach(MIMEText(html, "html", "utf-8"))
            msg.attach(MIMEText(render_template(body, row), "plain", "utf-8"))
        else:
            msg.attach(MIMEText(render_template(body, row), "plain", "utf-8"))
        return msg

    def _send_one(self, server, user, disp, subject, body, to_addr, row, as_html):
        msg = self._build_message(user, disp, subject, body, to_addr, row, as_html)
        server.sendmail(user, [to_addr], msg.as_string())

    def _send_test(self):
        try:
            cfg = self._collect_settings()
        except ValueError as e:
            messagebox.showwarning("信息不完整", str(e))
            return
        host, port, sec, user, pwd, disp, subject, body = cfg
        self._log("正在连接 %s:%s ..." % (host, port))
        try:
            server = self._connect(host, port, sec)
            self._send_one(server, user, disp, "[测试] " + subject, body, user,
                           {"姓名": "我"}, self.var_html.get())
            server.quit()
            self._log("√ 测试邮件已发送到 %s，请查收（可能在垃圾箱）" % user)
            messagebox.showinfo("成功", "测试邮件已发送到 %s，请查收。\n没收到的话翻一下垃圾邮件。" % user)
        except Exception as e:
            self._log("× 测试失败：%s" % e)
            messagebox.showerror("测试失败", "%s\n\n常见原因：\n1. 授权码不对（Gmail/Outlook 需要在账号安全设置里生成应用专用密码）\n2. 没开启 SMTP 服务\n3. 网络/防火墙拦截" % e)

    def _start_send(self):
        if self.sending:
            return
        if not self.rows:
            messagebox.showwarning("提示", "请先读取收件人名单（② 区）")
            return
        try:
            cfg = self._collect_settings()
        except ValueError as e:
            messagebox.showwarning("信息不完整", str(e))
            return
        if not messagebox.askyesno("确认发送",
                                   "即将向 %d 位收件人发送个性化邮件，确认开始？" % len(self.rows)):
            return
        self._save_config()
        self.stop_flag.clear()
        self.sending = True
        self.btn_send.config(state="disabled")
        self.btn_stop.config(state="normal")
        threading.Thread(target=self._send_worker, args=cfg, daemon=True).start()

    def _send_worker(self, host, port, sec, user, pwd, disp, subject, body):
        as_html = self.var_html.get()
        delay = max(0.0, float(self.spn_delay.get() or 0))
        ok, fail = 0, 0
        server = None
        try:
            self._log("连接 %s:%s ..." % (host, port))
            server = self._connect(host, port, sec)
            self._log("登录成功，开始发送 %d 封" % len(self.rows))
        except Exception as e:
            self._log("× 连接/登录失败：%s" % e)
            server = None

        for i, row in enumerate(self.rows, 1):
            if self.stop_flag.is_set():
                self._log("■ 已手动停止（第 %d 封之后）" % (i - 1))
                break
            to_addr = row.get("邮箱", "").strip()
            if not to_addr:
                continue
            try:
                if server is None:
                    server = self._connect(host, port, sec)
                self._send_one(server, user, disp, subject, body, to_addr, row, as_html)
                ok += 1
                self._log("[%d/%d] √ %s" % (i, len(self.rows), to_addr))
            except smtplib.SMTPServerDisconnected:
                try:
                    server = self._connect(host, port, sec)
                    self._send_one(server, user, disp, subject, body, to_addr, row, as_html)
                    ok += 1
                    self._log("[%d/%d] √ %s（重连后成功）" % (i, len(self.rows), to_addr))
                except Exception as e:
                    fail += 1
                    self._log("[%d/%d] × %s：%s" % (i, len(self.rows), to_addr, e))
            except Exception as e:
                fail += 1
                self._log("[%d/%d] × %s：%s" % (i, len(self.rows), to_addr, e))
            self.log_q.put(("progress", "进度 %d/%d（成功 %d，失败 %d）" % (i, len(self.rows), ok, fail)))
            if delay and i < len(self.rows):
                # 分片 sleep，便于及时响应停止
                waited = 0.0
                while waited < delay and not self.stop_flag.is_set():
                    time.sleep(0.2)
                    waited += 0.2

        if server:
            try:
                server.quit()
            except Exception:
                pass
        self.log_q.put(("done", "发送结束：成功 %d，失败 %d" % (ok, fail)))

    # ---------- 日志 ----------
    def _log(self, text):
        self.log_q.put(("log", text))

    def _poll_log(self):
        try:
            while True:
                kind, text = self.log_q.get_nowait()
                if kind == "progress":
                    self.lbl_progress.config(text=text)
                elif kind == "done":
                    self.lbl_progress.config(text=text)
                    self._log("— " + text)
                    self.sending = False
                    self.btn_send.config(state="normal")
                    self.btn_stop.config(state="disabled")
                else:
                    self.txt_log.insert("end", text + "\n")
                    self.txt_log.see("end")
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log)


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")
    except Exception:
        pass
    MailSenderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
