"""
HWP to PDF 일괄 변환기 (HWP to PDF Batch Converter)

한컴오피스(한글)의 COM 자동화 인터페이스(HWPFrame.HwpObject)를 이용하여
HWP/HWPX 문서를 PDF로 일괄 변환하는 윈도우 데스크탑 GUI 애플리케이션.

주요 특징:
  - 드래그 앤 드롭을 통한 파일/폴더 일괄 추가
  - 원본 폴더 구조를 유지한 채 PDF 저장
  - 백그라운드 스레드 변환으로 UI 프리징 방지
  - 변환 도중 중단(Stop) 기능
  - 파일별 실시간 상태 표시 (대기/변환 중/성공/실패/중지됨)
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pythoncom
import win32com.client as win32
from tkinterdnd2 import TkinterDnD, DND_FILES

# ──────────────────────────────────────────────
# 상수 정의
# ──────────────────────────────────────────────
BG_COLOR = "#f8f9fa"               # 앱 전체 배경색
BTN_CONVERT_COLOR = "#2563eb"       # 변환 버튼 기본 색상
BTN_STOP_COLOR = "#e11d48"          # 중지 버튼 색상
FONT_FAMILY = "Malgun Gothic"       # 기본 글꼴 (맑은 고딕)
HWP_EXTENSIONS = ('.hwp', '.hwpx')  # 지원하는 한글 문서 확장자
MAX_FAIL_DISPLAY = 5                # 결과 팝업에 표시할 최대 실패 사유 수

# ──────────────────────────────────────────────
# 전역 상태 변수
# ──────────────────────────────────────────────
is_converting = False   # 현재 변환 작업이 진행 중인지 여부
stop_requested = False  # 사용자가 중지 버튼을 눌렀는지 여부


# ──────────────────────────────────────────────
# UI 상태 업데이트 헬퍼 함수
# ──────────────────────────────────────────────
def _update_tree_status(item_id, status_text):
    """Treeview 특정 항목의 상태 컬럼을 안전하게 업데이트한다. (메인 스레드에서 실행)"""
    if item_id:
        root.after(0, lambda: tree.set(item_id, column="status", value=status_text))


def _set_buttons_enabled(enabled):
    """파일추가/저장경로/목록비우기 버튼의 활성 상태를 일괄 변경한다."""
    state = tk.NORMAL if enabled else tk.DISABLED
    btn_select.config(state=state)
    btn_output.config(state=state)
    btn_clear.config(state=state)


def _reset_convert_button():
    """변환 버튼을 초기(파란색, 변환 시작) 상태로 복원한다."""
    btn_convert.config(
        text="📄 PDF로 일괄 변환 시작",
        bg=BTN_CONVERT_COLOR,
        state=tk.NORMAL,
    )


# ──────────────────────────────────────────────
# 변환 완료 후 결과 표시 (메인 스레드)
# ──────────────────────────────────────────────
def _show_result(success_count, fail_count, stop_count, fail_reasons):
    """
    변환 작업이 끝난 뒤 호출되며, 결과 요약 팝업을 띄우고 UI를 초기 상태로 복원한다.
    반드시 메인(UI) 스레드에서 실행되어야 한다.
    """
    global is_converting
    is_converting = False
    _reset_convert_button()
    _set_buttons_enabled(True)
    progress['value'] = 0

    # 실패 사유 문자열 생성 (최대 MAX_FAIL_DISPLAY건)
    fail_msg = ""
    if fail_reasons:
        fail_msg = "\n".join(fail_reasons[:MAX_FAIL_DISPLAY])
        remaining = len(fail_reasons) - MAX_FAIL_DISPLAY
        if remaining > 0:
            fail_msg += f"\n... 외 {remaining}건"

    save_location = "지정한 폴더에 저장되었습니다." if output_dir else "원본 폴더에 저장되었습니다."

    # 상황별 메시지 박스 표시
    if stop_requested:
        messagebox.showinfo(
            "변환 중지",
            f"사용자에 의해 변환이 중단되었습니다.\n\n"
            f"🔹 성공: {success_count} 건\n🔹 실패: {fail_count} 건\n🔹 중단됨: {stop_count} 건",
        )
    elif fail_count == 0:
        messagebox.showinfo(
            "변환 완료",
            f"모든 파일 변환 성공!\n\n🔹 성공: {success_count} 건\n\n변환된 파일은 {save_location}",
        )
    elif success_count > 0:
        messagebox.showwarning(
            "변환 완료 (일부 실패)",
            f"변환이 완료되었으나 일부 파일이 실패했습니다.\n\n"
            f"🔹 성공: {success_count} 건\n🔹 실패: {fail_count} 건\n\n"
            f"[실패 사유]\n{fail_msg}\n\n성공한 파일은 {save_location}",
        )
    else:
        messagebox.showerror(
            "변환 실패",
            f"모든 파일 변환에 실패했습니다.\n\n🔹 실패: {fail_count} 건\n\n"
            f"[실패 사유]\n{fail_msg}\n\n문제가 계속되면 한글 프로그램 설정이나 라이선스 상태를 확인해 주세요.",
        )


def _show_hwp_error(error_msg):
    """한글 프로그램 초기화 실패 시 에러 팝업을 띄우고 UI를 복원한다."""
    global is_converting
    is_converting = False
    _reset_convert_button()
    _set_buttons_enabled(True)
    progress['value'] = 0
    messagebox.showerror("오류", error_msg)


# ──────────────────────────────────────────────
# 핵심 변환 로직 (백그라운드 스레드에서 실행)
# ──────────────────────────────────────────────
def _convert_worker(hwp_paths):
    """
    백그라운드 스레드에서 실행되는 실제 HWP → PDF 변환 워커 함수.

    COM 객체는 스레드별로 초기화해야 하므로 pythoncom.CoInitialize()를 호출하며,
    모든 UI 업데이트는 root.after()를 통해 메인 스레드로 위임한다.
    """
    global stop_requested

    if not hwp_paths:
        root.after(0, lambda: _show_result(0, 0, 0, []))
        return

    # ── COM 라이브러리 초기화 (스레드별 필수) ──
    pythoncom.CoInitialize()
    hwp = None

    try:
        # 한글 프로그램을 백그라운드로 실행하고, 보안/경고 팝업을 억제한다.
        hwp = win32.Dispatch("HWPFrame.HwpObject")
        hwp.XHwpWindows.Item(0).Visible = False                              # 한글 창 숨기기
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")      # 외부 접근 보안 팝업 우회
        hwp.SetMessageBoxMode(0x20)                                          # 내부 경고 팝업 자동 닫기
    except Exception as e:
        pythoncom.CoUninitialize()
        msg = (
            f"한컴오피스가 설치되어 있지 않거나 실행할 수 없습니다.\n"
            f"(주의: PC에 한글 프로그램이 설치되어 있어야 합니다.)\n\n상세 오류: {e}"
        )
        root.after(0, lambda: _show_hwp_error(msg))
        return

    total = len(hwp_paths)
    root.after(0, lambda: progress.config(maximum=total))

    # ── 변환 결과 카운터 ──
    success_count = 0
    fail_count = 0
    fail_reasons = []
    processed_index = 0  # 마지막으로 처리(시도)한 인덱스

    # ── 파일별 순차 변환 ──
    for i, hwp_path in enumerate(hwp_paths):
        processed_index = i

        # 사용자 중지 체크 (루프 시작 시)
        if stop_requested:
            break

        data = file_data.get(hwp_path)
        item_id = data["item_id"] if data else None
        rel_dir = data["rel_dir"] if data else ""

        # UI에 "변환 중..." 상태 표시
        _update_tree_status(item_id, "⏳ 변환 중...")

        try:
            abs_hwp_path = os.path.abspath(hwp_path)

            # 저장 대상 디렉토리 결정 (출력 폴더 지정 여부에 따라 분기)
            if output_dir:
                dest_dir = os.path.join(output_dir, rel_dir)
                os.makedirs(dest_dir, exist_ok=True)
            else:
                dest_dir = os.path.dirname(abs_hwp_path)

            # PDF 파일 경로 생성 (확장자만 .pdf로 교체)
            base_name = os.path.splitext(os.path.basename(abs_hwp_path))[0]
            pdf_path = os.path.join(dest_dir, base_name + ".pdf")

            # HWP 파일 열기
            hwp.Open(abs_hwp_path, "HWP", "forceopen:true")

            # 파일을 연 직후에도 중지 요청이 들어왔는지 확인
            if stop_requested:
                hwp.Run("FileClose")
                break

            # FileSaveAs_S 액션을 통해 PDF 포맷으로 저장
            hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
            hwp.HParameterSet.HFileOpenSave.filename = pdf_path
            hwp.HParameterSet.HFileOpenSave.Format = "PDF"
            hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)

            # 문서 닫기 (다음 파일 처리를 위해 반드시 닫아야 함)
            hwp.Run("FileClose")

            # 저장 결과 검증: 실제로 PDF 파일이 디스크에 생성되었는지 확인
            if os.path.exists(pdf_path):
                success_count += 1
                _update_tree_status(item_id, "✅ 성공")
            else:
                fail_count += 1
                reason = "PDF 파일이 생성되지 않았습니다."
                fail_reasons.append(f"- {os.path.basename(abs_hwp_path)}: {reason}")
                _update_tree_status(item_id, "❌ 실패")

        except Exception as e:
            fail_count += 1
            fail_reasons.append(f"- {os.path.basename(hwp_path)}: {e}")
            _update_tree_status(item_id, "❌ 실패")

        # 진행바 업데이트
        root.after(0, lambda val=i + 1: progress.config(value=val))

    # ── 한글 프로그램 종료 및 COM 해제 ──
    if hwp:
        hwp.Quit()
    pythoncom.CoUninitialize()

    # ── 중지된 경우: 아직 처리하지 못한 나머지 항목들을 "중지됨"으로 표시 ──
    stop_count = 0
    if stop_requested:
        for j in range(processed_index, total):
            data = file_data.get(hwp_paths[j])
            if data and data["item_id"]:
                _update_tree_status(data["item_id"], "⚠️ 중지됨")
                stop_count += 1

    # ── 메인 스레드에서 결과 팝업 표시 ──
    root.after(0, lambda: _show_result(success_count, fail_count, stop_count, fail_reasons))


# ──────────────────────────────────────────────
# 파일 목록 관리 함수
# ──────────────────────────────────────────────
def _add_to_list(filepath, rel_dir):
    """
    파일 경로를 Treeview 목록에 추가한다.
    이미 목록에 있는 파일은 중복 추가하지 않는다.

    Args:
        filepath: HWP/HWPX 파일의 절대 경로
        rel_dir:  출력 시 유지할 상대 디렉토리 경로 (예: "A폴더/하위폴더")
    """
    if filepath not in file_data:
        item_id = tree.insert("", "end", values=("대기", os.path.basename(filepath), filepath))
        file_data[filepath] = {"item_id": item_id, "rel_dir": rel_dir}


def _clear_list():
    """Treeview의 모든 항목과 내부 데이터를 초기화한다."""
    tree.delete(*tree.get_children())
    file_data.clear()


# ──────────────────────────────────────────────
# 사용자 이벤트 핸들러
# ──────────────────────────────────────────────
def _on_select_files():
    """파일 탐색기 대화상자를 열어 HWP/HWPX 파일을 선택·추가한다."""
    file_paths = filedialog.askopenfilenames(
        title="HWP 파일 선택",
        filetypes=[("HWP Files", "*.hwp *.hwpx")],
    )
    for path in file_paths:
        _add_to_list(path, "")


def _on_select_output_dir():
    """PDF 저장 대상 폴더를 사용자가 직접 지정할 수 있는 대화상자를 연다."""
    global output_dir
    folder = filedialog.askdirectory(title="PDF 저장 폴더 선택")
    if folder:
        output_dir = folder
        lbl_output.config(
            text=f"저장 폴더: {output_dir}\n(바탕화면 기본 / 여러 파일이나 폴더를 여기에 드래그 앤 드롭 하세요!)"
        )


def _on_start_conversion():
    """
    변환 시작/중지 토글 핸들러.
    - 변환 중이 아닐 때: 목록의 모든 파일에 대해 변환을 시작한다.
    - 변환 중일 때: 중지 플래그를 설정하여 워커 스레드가 안전하게 종료되도록 한다.
    """
    global is_converting, stop_requested

    if is_converting:
        # 중지 요청
        stop_requested = True
        btn_convert.config(text="🛑 중지 처리 중...", state=tk.DISABLED)
        return

    hwp_paths = list(file_data.keys())
    if not hwp_paths:
        messagebox.showwarning("경고", "변환할 HWP/HWPX 파일을 먼저 추가해주세요.")
        return

    # 상태 플래그 초기화
    is_converting = True
    stop_requested = False

    # 모든 항목의 상태를 "대기"로 리셋
    for path in hwp_paths:
        tree.set(file_data[path]["item_id"], column="status", value="대기")

    # UI를 "변환 중" 모드로 전환
    btn_convert.config(text="🛑 변환 중단하기", bg=BTN_STOP_COLOR)
    _set_buttons_enabled(False)

    # 백그라운드 스레드에서 변환 작업 시작 (UI 블로킹 방지)
    threading.Thread(target=_convert_worker, args=(hwp_paths,), daemon=True).start()


def _on_drop_files(event):
    """
    드래그 앤 드롭 이벤트 핸들러.
    - 폴더를 드롭하면 하위의 모든 HWP/HWPX 파일을 재귀적으로 탐색하여 추가한다.
      이때 드롭한 폴더 이름부터의 상대 경로를 기록하여 PDF 저장 시 동일 구조를 유지한다.
    - 개별 파일을 드롭하면 해당 파일만 추가한다.
    """
    if is_converting:
        return  # 변환 도중에는 새 파일 추가 방지

    paths = root.tk.splitlist(event.data)
    for path in paths:
        path = str(path)  # 경로를 문자열로 확정

        if os.path.isdir(path):
            # 폴더 드롭: 부모 디렉토리를 기준으로 상대 경로를 계산
            # 예) "C:/문서/A폴더" 드롭 시 base_dir = "C:/문서"
            #     → A폴더/하위/파일.hwp 의 rel_dir = "A폴더/하위"
            normalized = os.path.normpath(path)
            base_dir = os.path.dirname(normalized)

            for dir_path, _, filenames in os.walk(normalized):
                for filename in filenames:
                    if filename.lower().endswith(HWP_EXTENSIONS):
                        full_path = os.path.normpath(os.path.join(dir_path, filename))
                        rel_dir = os.path.relpath(os.path.dirname(full_path), base_dir)
                        if rel_dir == ".":
                            rel_dir = ""
                        _add_to_list(full_path, rel_dir)

        elif os.path.isfile(path) and path.lower().endswith(HWP_EXTENSIONS):
            _add_to_list(os.path.normpath(path), "")


# ══════════════════════════════════════════════
# GUI 레이아웃 구성
# ══════════════════════════════════════════════
root = TkinterDnD.Tk()
root.title("HWP to PDF 일괄 변환기")
root.geometry("750x650")
root.resizable(True, True)
root.minsize(500, 250)
root.config(bg=BG_COLOR)

# ── 테마 및 스타일 ──
style = ttk.Style()
style.theme_use("clam")
style.configure("Treeview", font=(FONT_FAMILY, 9), rowheight=25)
style.configure("Treeview.Heading", font=(FONT_FAMILY, 10, "bold"))

# ── 전역 데이터 ──
file_data = {}  # { 파일절대경로: {"item_id": Treeview 행 ID, "rel_dir": 상대 디렉토리} }
output_dir = os.path.join(os.path.expanduser("~"), "Desktop")  # 기본 저장 경로: 바탕화면

# ── 메인 프레임 ──
main_frame = tk.Frame(root, padx=20, pady=20, bg=BG_COLOR)
main_frame.pack(fill=tk.BOTH, expand=True)

# ── 타이틀 & 안내 문구 ──
tk.Label(
    main_frame,
    text="HWP / HWPX 파일 일괄 PDF 변환기",
    font=(FONT_FAMILY, 16, "bold"),
    bg=BG_COLOR,
    fg="#333333",
).pack(pady=(0, 5))

tk.Label(
    main_frame,
    text="※ 안내: PC에 한컴오피스(한글)가 설치되어 있어야 정상 작동합니다.",
    font=(FONT_FAMILY, 9),
    bg=BG_COLOR,
    fg="#e11d48",
).pack(pady=(0, 15))

# ── 상단 버튼 영역 ──
top_btn_frame = tk.Frame(main_frame, bg=BG_COLOR)
top_btn_frame.pack(fill=tk.X, pady=(0, 10))

btn_select = tk.Button(
    top_btn_frame, text="📁 파일 추가", command=_on_select_files,
    width=15, font=(FONT_FAMILY, 10), bg="#ffffff", relief="solid", bd=1,
)
btn_select.pack(side=tk.LEFT, padx=(0, 10))

btn_output = tk.Button(
    top_btn_frame, text="📂 저장 경로 지정", command=_on_select_output_dir,
    width=15, font=(FONT_FAMILY, 10), bg="#ffffff", relief="solid", bd=1,
)
btn_output.pack(side=tk.LEFT)

btn_clear = tk.Button(
    top_btn_frame, text="🗑️ 목록 비우기", command=_clear_list,
    width=15, font=(FONT_FAMILY, 10), bg="#ffffff", relief="solid", bd=1,
)
btn_clear.pack(side=tk.RIGHT)

# ── 저장 폴더 안내 레이블 ──
lbl_output = tk.Label(
    main_frame,
    text=f"저장 폴더: {output_dir}\n(바탕화면 기본 / 여러 파일이나 폴더를 여기에 드래그 앤 드롭 하세요!)",
    font=(FONT_FAMILY, 9),
    bg=BG_COLOR,
    fg="#0284c7",
    justify="left",
)
lbl_output.pack(anchor="w", pady=(0, 5))

# ── 파일 목록 Treeview ──
tree_frame = tk.Frame(main_frame)
tree_frame.pack(fill=tk.BOTH, expand=True)

scrollbar = ttk.Scrollbar(tree_frame)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

columns = ("status", "filename", "folder")
tree = ttk.Treeview(
    tree_frame, columns=columns, show="headings",
    yscrollcommand=scrollbar.set, selectmode="extended",
)
tree.heading("status", text="상태")
tree.column("status", width=80, anchor="center")
tree.heading("filename", text="파일명")
tree.column("filename", width=250, anchor="w")
tree.heading("folder", text="파일경로")
tree.column("folder", width=350, anchor="w")
tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar.config(command=tree.yview)

# ── 드래그 앤 드롭 바인딩 ──
tree.drop_target_register(DND_FILES)
tree.dnd_bind("<<Drop>>", _on_drop_files)

# ── 진행률 바 ──
progress = ttk.Progressbar(main_frame, orient="horizontal", mode="determinate")
progress.pack(pady=(15, 10), fill=tk.X)

# ── 변환 시작/중지 버튼 ──
btn_convert = tk.Button(
    main_frame, text="📄 PDF로 일괄 변환 시작", command=_on_start_conversion,
    height=2, bg=BTN_CONVERT_COLOR, fg="white",
    font=(FONT_FAMILY, 12, "bold"), relief="flat", cursor="hand2",
)
btn_convert.pack(fill=tk.X)

# ── 메인 이벤트 루프 시작 ──
root.mainloop()
