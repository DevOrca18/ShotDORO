import cv2
import numpy as np
import pytesseract
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from PIL import Image, ImageTk
import csv
import datetime
import time
import threading
import re
import os
import pandas as pd


class VideoAmmoAnalyzer:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ShotDORO 영상 분석기")
        self.root.geometry("900x700")

        # 상태 변수들
        self.video_path = None
        self.image_path = None  # 추가: 영역 설정용 이미지 경로
        self.total_ammo_region = None
        self.current_ammo_region = None
        self.shot_data = []
        self.video_fps = 30
        self.analysis_running = False
        self.existing_csv_path = None  # 추가: 기존 CSV 경로

        # GUI 설정
        self.setup_gui()

        # Tesseract 경로 설정 (필요시 수정)
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

    def setup_gui(self):
        # 메인 프레임
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 비디오 프레임
        video_frame = tk.Frame(main_frame)
        video_frame.pack(fill=tk.BOTH, expand=True)

        self.video_label = tk.Label(video_frame, text="영상을 로드해주세요", bg="gray")
        self.video_label.pack(fill=tk.BOTH, expand=True)

        # 컨트롤 프레임
        control_frame = tk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=10)

        # 첫 번째 행 버튼들
        btn_frame1 = tk.Frame(control_frame)
        btn_frame1.pack(fill=tk.X, pady=2)

        tk.Button(btn_frame1, text="영상 로드", command=self.load_video, bg="lightblue").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame1, text="이미지 로드", command=self.load_image, bg="lightcyan").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame1, text="총 장탄수 영역", command=self.set_total_ammo_region, bg="lightgreen").pack(side=tk.LEFT,
                                                                                                         padx=5)
        tk.Button(btn_frame1, text="현재 총알 영역", command=self.set_current_ammo_region, bg="lightgreen").pack(side=tk.LEFT,
                                                                                                           padx=5)

        # 두 번째 행 버튼들
        btn_frame2 = tk.Frame(control_frame)
        btn_frame2.pack(fill=tk.X, pady=2)

        tk.Button(btn_frame2, text="🚀 빠른 분석 시작", command=self.start_analysis, bg="orange",
                  font=("Arial", 12, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame2, text="분석 중지", command=self.stop_analysis, bg="red").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame2, text="CSV 저장", command=self.save_csv, bg="yellow").pack(side=tk.LEFT, padx=5)

        # 추가: CSV 등록 버튼
        tk.Button(btn_frame2, text="📋 CSV 등록", command=self.load_existing_csv, bg="lightcoral").pack(side=tk.LEFT,
                                                                                                     padx=5)
        tk.Button(btn_frame2, text="💾 사격시간 추가", command=self.add_shot_times_to_csv, bg="lightseagreen").pack(
            side=tk.LEFT, padx=5)

        # 설정 프레임
        setting_frame = tk.Frame(control_frame)
        setting_frame.pack(fill=tk.X, pady=5)

        tk.Label(setting_frame, text="프레임 건너뛰기:").pack(side=tk.LEFT)
        self.skip_frames = tk.IntVar(value=1)
        tk.Spinbox(setting_frame, from_=1, to=10, textvariable=self.skip_frames, width=5).pack(side=tk.LEFT, padx=5)
        tk.Label(setting_frame, text="(1=모든 프레임, 5=5프레임마다)").pack(side=tk.LEFT, padx=10)

        # 알림 설정
        self.sound_alert = tk.BooleanVar(value=True)
        tk.Checkbutton(setting_frame, text="🔊 사격감지 알림", variable=self.sound_alert).pack(side=tk.LEFT, padx=10)

        # 시간 오차 범위 설정 추가
        tk.Label(setting_frame, text="시간 오차(초):").pack(side=tk.LEFT, padx=(20, 0))
        self.time_tolerance = tk.DoubleVar(value=0.1)
        tk.Spinbox(setting_frame, from_=0.01, to=1.0, increment=0.01, textvariable=self.time_tolerance, width=6).pack(
            side=tk.LEFT, padx=5)

        # 진행 상황 프레임
        progress_frame = tk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=5)

        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=2)

        # 상태 표시
        status_frame = tk.Frame(main_frame)
        status_frame.pack(fill=tk.X)

        self.status_label = tk.Label(status_frame, text="상태: 대기 중", font=("Arial", 12))
        self.status_label.pack(side=tk.LEFT)

        self.info_label = tk.Label(status_frame, text="", font=("Arial", 10))
        self.info_label.pack(side=tk.RIGHT)

        # 실시간 탄약 정보 표시
        ammo_info_frame = tk.Frame(main_frame)
        ammo_info_frame.pack(fill=tk.X, pady=2)

        self.current_ammo_label = tk.Label(ammo_info_frame, text="현재 탄약: -",
                                           font=("Arial", 14, "bold"), fg="blue")
        self.current_ammo_label.pack(side=tk.LEFT)

        self.total_ammo_label = tk.Label(ammo_info_frame, text="총 장탄: -",
                                         font=("Arial", 14, "bold"), fg="green")
        self.total_ammo_label.pack(side=tk.LEFT, padx=20)

        self.shot_count_label = tk.Label(ammo_info_frame, text="감지된 사격: 0발",
                                         font=("Arial", 14, "bold"), fg="red")
        self.shot_count_label.pack(side=tk.RIGHT)

        # CSV 정보 표시 프레임 추가
        csv_info_frame = tk.Frame(main_frame)
        csv_info_frame.pack(fill=tk.X, pady=2)

        self.csv_info_label = tk.Label(csv_info_frame, text="등록된 CSV: 없음",
                                       font=("Arial", 11), fg="purple")
        self.csv_info_label.pack(side=tk.LEFT)

        # 결과 표시 프레임
        result_frame = tk.Frame(main_frame)
        result_frame.pack(fill=tk.X, pady=5)

        tk.Label(result_frame, text="분석 결과:", font=("Arial", 12, "bold")).pack(anchor=tk.W)

        # 스크롤 가능한 텍스트 위젯
        text_frame = tk.Frame(result_frame)
        text_frame.pack(fill=tk.X)

        self.result_text = tk.Text(text_frame, height=8, font=("Consolas", 9))
        scrollbar = tk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        self.result_text.config(yscrollcommand=scrollbar.set)

        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def load_existing_csv(self):
        """기존 CSV 파일 로드"""
        file_path = filedialog.askopenfilename(
            title="기존 CSV 파일 선택",
            filetypes=[("CSV 파일", "*.csv")]
        )

        if file_path:
            try:
                # CSV 파일 로드하여 구조 확인
                df = pd.read_csv(file_path, encoding='utf-8-sig')
                self.existing_csv_path = file_path

                # 시간 컬럼 확인 (f열 = 5번 인덱스, 또는 'time' 컬럼)
                time_column = None
                if len(df.columns) > 5:  # f열 확인
                    time_column = df.columns[5]
                elif 'time' in df.columns:
                    time_column = 'time'
                else:
                    # 시간 같은 컬럼 찾기
                    for col in df.columns:
                        if 'time' in col.lower():
                            time_column = col
                            break

                if time_column:
                    self.csv_info_label.config(
                        text=f"등록된 CSV: {os.path.basename(file_path)} (행: {len(df)}, 시간열: {time_column})",
                        fg="green"
                    )
                    self.result_text.insert(tk.END, f"✅ CSV 등록 완료: {len(df)}행, 시간 컬럼: '{time_column}'\n")
                else:
                    messagebox.showwarning("경고", "시간 정보를 찾을 수 없습니다. 'time' 컬럼이나 f열을 확인해주세요.")
                    self.existing_csv_path = None

            except Exception as e:
                messagebox.showerror("오류", f"CSV 파일을 읽을 수 없습니다: {e}")
                self.existing_csv_path = None

    def add_shot_times_to_csv(self):
        """기존 CSV에 사격 시간 추가"""
        if not self.existing_csv_path:
            messagebox.showwarning("경고", "먼저 CSV 파일을 등록해주세요.")
            return

        if not self.shot_data:
            messagebox.showwarning("경고", "사격 분석 데이터가 없습니다. 먼저 영상 분석을 실행해주세요.")
            return

        try:
            # 기존 CSV 로드
            df = pd.read_csv(self.existing_csv_path, encoding='utf-8-sig')

            # 시간 컬럼 찾기
            time_column = None
            if len(df.columns) > 5:  # f열
                time_column = df.columns[5]
            elif 'time' in df.columns:
                time_column = 'time'
            else:
                for col in df.columns:
                    if 'time' in col.lower():
                        time_column = col
                        break

            if not time_column:
                messagebox.showerror("오류", "시간 컬럼을 찾을 수 없습니다.")
                return

            # 사격시간 컬럼 추가 (이미 있으면 초기화)
            df['사격시간'] = ''

            # 사격 데이터와 매칭
            tolerance = self.time_tolerance.get()
            matched_count = 0

            for shot in self.shot_data:
                shot_time = shot['time_seconds']

                # 시간 오차 범위 내의 행 찾기
                time_diff = abs(df[time_column] - shot_time)
                matching_rows = time_diff <= tolerance

                if matching_rows.any():
                    df.loc[matching_rows, '사격시간'] = 'O'
                    matched_count += len(df[matching_rows])

            # 결과 저장
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = os.path.splitext(os.path.basename(self.existing_csv_path))[0]
            default_name = f"{base_name}_with_shots_{timestamp}.csv"

            save_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV 파일", "*.csv")],
                title="사격시간이 추가된 CSV 저장",
                initialvalue=default_name
            )

            if save_path:
                df.to_csv(save_path, index=False, encoding='utf-8-sig')

                # 결과 메시지
                summary = f"\n{'=' * 50}\n"
                summary += f"📊 사격시간 추가 완료!\n"
                summary += f"📁 원본 CSV: {os.path.basename(self.existing_csv_path)} ({len(df)}행)\n"
                summary += f"🎯 감지된 사격: {len(self.shot_data)}발\n"
                summary += f"✅ 매칭된 행: {matched_count}개\n"
                summary += f"⏰ 시간 오차 범위: ±{tolerance}초\n"
                summary += f"💾 저장 경로: {save_path}\n"
                summary += f"{'=' * 50}\n"

                self.result_text.insert(tk.END, summary)
                self.result_text.see(tk.END)

                messagebox.showinfo("완료",
                                    f"사격시간이 추가된 CSV가 저장되었습니다!\n\n"
                                    f"감지된 사격: {len(self.shot_data)}발\n"
                                    f"매칭된 행: {matched_count}개\n"
                                    f"저장 위치: {os.path.basename(save_path)}")

        except Exception as e:
            messagebox.showerror("오류", f"CSV 처리 중 오류가 발생했습니다: {e}")

    def load_video(self):
        file_path = filedialog.askopenfilename(
            title="영상 파일 선택",
            filetypes=[("비디오 파일", "*.mp4 *.avi *.mov *.mkv *.wmv")]
        )
        if file_path:
            self.video_path = file_path
            self.load_first_frame()

    def load_image(self):
        """영역 설정용 이미지 로드"""
        file_path = filedialog.askopenfilename(
            title="영역 설정용 이미지 선택",
            filetypes=[("이미지 파일", "*.png *.jpg *.jpeg *.bmp *.tiff")]
        )
        if file_path:
            try:
                # 이미지 로드 테스트
                test_frame = cv2.imread(file_path)
                if test_frame is None:
                    messagebox.showerror("오류", "이미지를 로드할 수 없습니다.")
                    return

                self.image_path = file_path

                # 이미지 정보 표시
                height, width = test_frame.shape[:2]
                file_name = os.path.basename(file_path)
                self.status_label.config(text=f"이미지 로드됨: {file_name} ({width}x{height})")

                messagebox.showinfo("완료", f"이미지가 로드되었습니다!\n{file_name}\n크기: {width}x{height}")

            except Exception as e:
                messagebox.showerror("오류", f"이미지 로드 중 오류: {e}")
                self.image_path = None

    def load_first_frame(self):
        cap = cv2.VideoCapture(self.video_path)
        ret, frame = cap.read()
        if ret:
            self.video_fps = cap.get(cv2.CAP_PROP_FPS)
            total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
            duration = total_frames / self.video_fps

            # 첫 프레임 표시
            self.display_frame(frame)

            self.info_label.config(text=f"FPS: {self.video_fps:.1f} | 프레임: {total_frames} | 길이: {duration:.1f}초")
            self.status_label.config(text="상태: 영상 로드됨")
        cap.release()

    def display_frame(self, frame):
        # 프레임을 GUI에 맞게 리사이즈
        height, width = frame.shape[:2]
        max_width, max_height = 800, 400

        scale = min(max_width / width, max_height / height)
        new_width = int(width * scale)
        new_height = int(height * scale)

        resized_frame = cv2.resize(frame, (new_width, new_height))
        rgb_frame = cv2.cvtColor(resized_frame, cv2.COLOR_BGR2RGB)
        pil_image = Image.fromarray(rgb_frame)
        photo = ImageTk.PhotoImage(pil_image)

        self.video_label.config(image=photo, text="")
        self.video_label.image = photo
        self.current_frame = frame  # 원본 프레임 저장

    def set_total_ammo_region(self):
        self.total_ammo_region = self.select_region("총 장탄수 영역을 선택하세요")

    def set_current_ammo_region(self):
        self.current_ammo_region = self.select_region("현재 총알 수 영역을 선택하세요")

    def select_region(self, title):
        # img.png 파일 찾기
        img_path = "img.png"
        if not os.path.exists(img_path):
            # 현재 디렉토리에 img.png가 없으면 파일 선택
            img_path = filedialog.askopenfilename(
                title="샘플 이미지 선택 (img.png)",
                filetypes=[("이미지 파일", "*.png *.jpg *.jpeg")],
                initialfile="img.png"
            )
            if not img_path:
                messagebox.showwarning("경고", "샘플 이미지가 필요합니다.")
                return None

        try:
            # img.png 파일 로드
            frame = cv2.imread(img_path)
            if frame is None:
                messagebox.showerror("오류", f"이미지를 로드할 수 없습니다: {img_path}")
                return None
        except Exception as e:
            messagebox.showerror("오류", f"이미지 로드 중 오류: {e}")
            return None

        # 새 창에서 영역 선택
        region_window = tk.Toplevel(self.root)
        region_window.title(f"{title} - img.png")
        region_window.geometry("900x700")

        image_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        pil_image = Image.fromarray(image_rgb)

        # 화면에 맞게 스케일링
        original_size = pil_image.size
        pil_image.thumbnail((800, 600), Image.Resampling.LANCZOS)
        display_size = pil_image.size

        photo = ImageTk.PhotoImage(pil_image)

        canvas = tk.Canvas(region_window, width=display_size[0], height=display_size[1])
        canvas.pack(pady=20)
        canvas.create_image(display_size[0] // 2, display_size[1] // 2, image=photo)
        canvas.image = photo

        # 선택 변수
        selection = {'start_x': 0, 'start_y': 0, 'end_x': 0, 'end_y': 0, 'rect_id': None, 'confirmed': False}

        def on_click(event):
            selection['start_x'], selection['start_y'] = event.x, event.y
            if selection['rect_id']:
                canvas.delete(selection['rect_id'])

        def on_drag(event):
            selection['end_x'], selection['end_y'] = event.x, event.y
            if selection['rect_id']:
                canvas.delete(selection['rect_id'])
            selection['rect_id'] = canvas.create_rectangle(
                selection['start_x'], selection['start_y'],
                selection['end_x'], selection['end_y'],
                outline="red", width=3
            )

        def confirm_selection():
            if abs(selection['end_x'] - selection['start_x']) > 10 and abs(
                    selection['end_y'] - selection['start_y']) > 10:
                selection['confirmed'] = True
                region_window.destroy()
            else:
                messagebox.showwarning("경고", "유효한 영역을 선택해주세요.")

        canvas.bind("<Button-1>", on_click)
        canvas.bind("<B1-Motion>", on_drag)

        confirm_btn = tk.Button(region_window, text="✓ 선택 확인", command=confirm_selection,
                                bg="green", fg="white", font=("Arial", 12))
        confirm_btn.pack(pady=10)

        # 안내 텍스트 추가
        info_label = tk.Label(region_window, text="마우스로 드래그하여 영역을 선택하세요",
                              font=("Arial", 10), fg="blue")
        info_label.pack()

        region_window.wait_window()

        if selection['confirmed']:
            # 실제 이미지 좌표로 변환 (img.png와 영상이 같은 해상도)
            scale_x = original_size[0] / display_size[0]
            scale_y = original_size[1] / display_size[1]

            x1 = int(min(selection['start_x'], selection['end_x']) * scale_x)
            y1 = int(min(selection['start_y'], selection['end_y']) * scale_y)
            x2 = int(max(selection['start_x'], selection['end_x']) * scale_x)
            y2 = int(max(selection['start_y'], selection['end_y']) * scale_y)

            return (x1, y1, x2, y2)
        return None

    def play_alert_sound(self):
        """사격 감지 시 알림음 재생"""
        try:
            # 윈도우 시스템 사운드 재생
            import winsound
            winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
        except:
            # 시스템 벨 사운드
            try:
                import os
                os.system("echo \a")
            except:
                pass

    def extract_number_from_region(self, frame, region):
        if region is None:
            return None

        x1, y1, x2, y2 = region
        roi = frame[y1:y2, x1:x2]

        if roi.size == 0:
            return None

        # 이미지 전처리 - 더 강화된 버전
        gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)

        # 크기 확대로 OCR 정확도 향상
        if roi.shape[0] < 50 or roi.shape[1] < 50:
            gray = cv2.resize(gray, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)

        # 노이즈 제거 및 이진화
        gray = cv2.medianBlur(gray, 3)
        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

        # 형태학적 연산으로 텍스트 정리
        kernel = np.ones((2, 2), np.uint8)
        thresh = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)
        thresh = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel)

        # OCR 설정 최적화
        custom_config = r'--oem 3 --psm 8 -c tessedit_char_whitelist=0123456789'

        try:
            text = pytesseract.image_to_string(thresh, config=custom_config)
            # 숫자만 추출
            numbers = re.findall(r'\d+', text.strip())
            if numbers:
                return int(numbers[0])
        except:
            pass

        return None

    def analyze_video(self):
        if not self.current_ammo_region:
            messagebox.showwarning("경고", "현재 총알 수 영역을 먼저 설정해주세요.")
            return

        cap = cv2.VideoCapture(self.video_path)
        total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
        fps = cap.get(cv2.CAP_PROP_FPS)

        self.shot_data = []
        previous_ammo = None
        frame_count = 0
        skip_frames = self.skip_frames.get()

        self.progress_bar['maximum'] = total_frames

        start_time = time.time()

        while self.analysis_running and cap.isOpened():
            ret, frame = cap.read()
            if not ret:
                break

            frame_count += 1

            # 프레임 건너뛰기
            if frame_count % skip_frames != 0:
                continue

            # 진행률 업데이트
            self.progress_bar['value'] = frame_count

            # 현재 탄약 수 추출
            current_ammo = self.extract_number_from_region(frame, self.current_ammo_region)
            total_ammo = self.extract_number_from_region(frame, self.total_ammo_region)

            # 실시간 탄약 정보 업데이트
            self.root.after(0, self.update_ammo_display, total_ammo, current_ammo, frame_count)

            # 사격 감지
            if current_ammo is not None and previous_ammo is not None:
                if current_ammo < previous_ammo:
                    # 사격 감지!
                    time_seconds = frame_count / fps
                    shot_time = f"{int(time_seconds // 60):02d}:{time_seconds % 60:06.3f}"

                    self.shot_data.append({
                        'frame': frame_count,
                        'time': shot_time,
                        'time_seconds': round(time_seconds, 3),
                        'total_ammo': total_ammo if total_ammo else '-',
                        'current_ammo': current_ammo,
                        'previous_ammo': previous_ammo,
                        'shots_fired': previous_ammo - current_ammo
                    })

                    # 🎯 사격 감지 알림 및 결과 표시
                    result = f"🎯 {len(self.shot_data):2d}. {shot_time} | {previous_ammo}→{current_ammo} ({previous_ammo - current_ammo}발)\n"
                    self.result_text.insert(tk.END, result)
                    self.result_text.see(tk.END)

                    # 알림음 재생
                    if self.sound_alert.get():
                        self.play_alert_sound()

                    # 사격 횟수 업데이트
                    self.root.after(0, self.update_shot_count, len(self.shot_data))

                    self.root.update()
            else:
                # 프레임별 탄약 정보만 표시 (사격 없을 때)
                if current_ammo is not None or total_ammo is not None:
                    info = f"프레임 {frame_count:5d} | 총탄:{total_ammo if total_ammo else '?'} 현재:{current_ammo if current_ammo else '?'}\n"
                    # 너무 많은 로그를 방지하기 위해 10프레임마다만 표시
                    if frame_count % (skip_frames * 10) == 0:
                        self.result_text.insert(tk.END, info)
                        self.result_text.see(tk.END)

            if current_ammo is not None:
                previous_ammo = current_ammo

            # 상태 업데이트 (100프레임마다)
            if frame_count % 100 == 0:
                elapsed = time.time() - start_time
                fps_actual = frame_count / elapsed if elapsed > 0 else 0
                self.status_label.config(
                    text=f"분석 중... {frame_count}/{total_frames} ({fps_actual:.1f} FPS)"
                )
                self.root.update()

        cap.release()

        if self.analysis_running:
            elapsed = time.time() - start_time
            self.status_label.config(
                text=f"분석 완료! {len(self.shot_data)}발 감지 (소요시간: {elapsed:.1f}초)"
            )

            # 최종 요약
            summary = f"\n{'=' * 50}\n"
            summary += f"📊 분석 완료: 총 {len(self.shot_data)}발의 사격 감지\n"
            summary += f"⏱️ 분석 시간: {elapsed:.1f}초\n"
            summary += f"🎬 총 프레임: {total_frames} (FPS: {fps:.1f})\n"
            summary += f"{'=' * 50}\n"
            self.result_text.insert(tk.END, summary)

        self.analysis_running = False

    def update_ammo_display(self, total_ammo, current_ammo, frame_count):
        """실시간 탄약 정보 표시 업데이트"""
        total_str = str(total_ammo) if total_ammo is not None else "?"
        current_str = str(current_ammo) if current_ammo is not None else "?"

        self.current_ammo_label.config(text=f"현재 탄약: {current_str}")
        self.total_ammo_label.config(text=f"총 장탄: {total_str}")

    def update_shot_count(self, count):
        """사격 횟수 업데이트"""
        self.shot_count_label.config(text=f"감지된 사격: {count}발")

    def start_analysis(self):
        if not self.video_path:
            messagebox.showwarning("경고", "영상 파일을 먼저 선택해주세요.")
            return

        if not self.current_ammo_region:
            messagebox.showwarning("경고", "현재 총알 수 영역을 설정해주세요.")
            return

        self.analysis_running = True
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "🚀 분석 시작...\n\n")

        # 백그라운드 스레드에서 분석 실행
        self.analysis_thread = threading.Thread(target=self.analyze_video, daemon=True)
        self.analysis_thread.start()

    def stop_analysis(self):
        self.analysis_running = False
        self.status_label.config(text="상태: 분석 중지됨")

    def save_csv(self):
        if not self.shot_data:
            messagebox.showinfo("정보", "저장할 데이터가 없습니다.")
            return

        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"shot_analysis_{timestamp}.csv"

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV 파일", "*.csv")],
            title="CSV 파일 저장",
            initialvalue=default_name
        )

        if file_path:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                fieldnames = ['shot_number', 'frame', 'time', 'time_seconds', 'total_ammo', 'current_ammo',
                              'previous_ammo', 'shots_fired']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

                writer.writeheader()
                for i, row in enumerate(self.shot_data, 1):
                    row['shot_number'] = i
                    writer.writerow(row)

            messagebox.showinfo("성공", f"📁 {len(self.shot_data)}개 데이터가 저장되었습니다!\n{file_path}")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    # 필요한 라이브러리 설치 안내
    try:
        import cv2
        import pytesseract
        import pandas as pd
    except ImportError as e:
        print("필요한 라이브러리를 설치해주세요:")
        print("pip install opencv-python pytesseract pillow pandas")
        print("또한 Tesseract OCR을 설치해야 합니다:")
        print("https://github.com/tesseract-ocr/tesseract")
        exit(1)

    app = VideoAmmoAnalyzer()
    app.run()