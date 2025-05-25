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
import random


class VideoAmmoAnalyzer:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ShotDORO Video Analyzer")
        self.root.geometry("1000x800")

        # State variables
        self.video_path = None
        self.sample_frames = []  # Store sample frames
        self.current_frame = None  # Selected frame (for region setup)
        self.total_ammo_region = None
        self.current_ammo_region = None
        self.shot_data = []
        self.video_fps = 30
        self.analysis_running = False
        self.existing_csv_path = None

        # Setup GUI
        self.setup_gui()

        # Set Tesseract path (modify if needed)
        try:
            pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        except:
            pass

    def setup_gui(self):
        # Main frame
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Video frame
        video_frame = tk.Frame(main_frame)
        video_frame.pack(fill=tk.BOTH, expand=True)

        self.video_label = tk.Label(video_frame, text="Please load a video", bg="gray")
        self.video_label.pack(fill=tk.BOTH, expand=True)

        # Control frame
        control_frame = tk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=10)

        # First row buttons
        btn_frame1 = tk.Frame(control_frame)
        btn_frame1.pack(fill=tk.X, pady=2)

        tk.Button(btn_frame1, text="🎬 Load Video", command=self.load_video, bg="lightblue",
                  font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame1, text="🖼️ Select Frame", command=self.select_frame, bg="lightcyan").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame1, text="🎯 Total Ammo Region", command=self.set_total_ammo_region, bg="lightgreen").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame1, text="🔫 Current Ammo Region", command=self.set_current_ammo_region, bg="lightgreen").pack(side=tk.LEFT, padx=5)

        # Second row buttons
        btn_frame2 = tk.Frame(control_frame)
        btn_frame2.pack(fill=tk.X, pady=2)

        tk.Button(btn_frame2, text="🚀 Start Analysis", command=self.start_analysis, bg="orange",
                  font=("Arial", 12, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame2, text="⏹️ Stop Analysis", command=self.stop_analysis, bg="red").pack(side=tk.LEFT, padx=5)

        # Third row buttons (CSV related)
        btn_frame3 = tk.Frame(control_frame)
        btn_frame3.pack(fill=tk.X, pady=2)

        tk.Button(btn_frame3, text="📋 Load CSV (before adding pings)", command=self.load_existing_csv, bg="lightcoral").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame3, text="➕ Add Shot Pings", command=self.add_shot_times_to_csv, bg="lightseagreen").pack(side=tk.LEFT, padx=5)

        # Settings frame
        setting_frame = tk.Frame(control_frame)
        setting_frame.pack(fill=tk.X, pady=5)

        tk.Label(setting_frame, text="Skip Frames:").pack(side=tk.LEFT)
        self.skip_frames = tk.IntVar(value=1)
        tk.Spinbox(setting_frame, from_=1, to=10, textvariable=self.skip_frames, width=5).pack(side=tk.LEFT, padx=5)

        # Alert settings
        self.sound_alert = tk.BooleanVar(value=True)
        tk.Checkbutton(setting_frame, text="🔊 Shot Detection Alert", variable=self.sound_alert).pack(side=tk.LEFT, padx=10)

        # Time tolerance setting
        tk.Label(setting_frame, text="Time Tolerance(sec):").pack(side=tk.LEFT, padx=(20, 0))
        self.time_tolerance = tk.DoubleVar(value=0.1)
        tk.Spinbox(setting_frame, from_=0.01, to=1.0, increment=0.01, textvariable=self.time_tolerance, width=6).pack(side=tk.LEFT, padx=5)

        # Progress frame
        progress_frame = tk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=5)

        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=2)

        # Status display
        status_frame = tk.Frame(main_frame)
        status_frame.pack(fill=tk.X)

        self.status_label = tk.Label(status_frame, text="Status: Waiting", font=("Arial", 12))
        self.status_label.pack(side=tk.LEFT)

        self.info_label = tk.Label(status_frame, text="", font=("Arial", 10))
        self.info_label.pack(side=tk.RIGHT)

        # Real-time ammo info display
        ammo_info_frame = tk.Frame(main_frame)
        ammo_info_frame.pack(fill=tk.X, pady=2)

        self.current_ammo_label = tk.Label(ammo_info_frame, text="Current Ammo: -",
                                           font=("Arial", 14, "bold"), fg="blue")
        self.current_ammo_label.pack(side=tk.LEFT)

        self.total_ammo_label = tk.Label(ammo_info_frame, text="Total Ammo: -",
                                         font=("Arial", 14, "bold"), fg="green")
        self.total_ammo_label.pack(side=tk.LEFT, padx=20)

        self.shot_count_label = tk.Label(ammo_info_frame, text="Detected Shots: 0",
                                         font=("Arial", 14, "bold"), fg="red")
        self.shot_count_label.pack(side=tk.RIGHT)

        # CSV and frame info display
        info_frame = tk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=2)

        self.csv_info_label = tk.Label(info_frame, text="Loaded CSV: None",
                                       font=("Arial", 11), fg="purple")
        self.csv_info_label.pack(side=tk.LEFT)

        self.frame_info_label = tk.Label(info_frame, text="Selected Frame: None",
                                         font=("Arial", 11), fg="darkorange")
        self.frame_info_label.pack(side=tk.RIGHT)

        # Results display frame
        result_frame = tk.Frame(main_frame)
        result_frame.pack(fill=tk.X, pady=5)

        tk.Label(result_frame, text="Analysis Results:", font=("Arial", 12, "bold")).pack(anchor=tk.W)

        # Scrollable text widget
        text_frame = tk.Frame(result_frame)
        text_frame.pack(fill=tk.X)

        self.result_text = tk.Text(text_frame, height=8, font=("Consolas", 9))
        scrollbar = tk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        self.result_text.config(yscrollcommand=scrollbar.set)

        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def load_existing_csv(self):
        """Load existing CSV file - error correction and usability improvements"""
        file_path = filedialog.askopenfilename(
            title="Select existing CSV file",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )

        if file_path:
            try:
                # Try CSV file with multiple encodings
                df = None
                encodings = ['utf-8-sig', 'utf-8', 'cp949', 'euc-kr', 'latin1']

                for encoding in encodings:
                    try:
                        df = pd.read_csv(file_path, encoding=encoding)
                        break
                    except:
                        continue

                if df is None:
                    messagebox.showerror("Error", "Cannot read CSV file. Please check the file format.")
                    return

                self.existing_csv_path = file_path

                # Improved automatic time column detection
                time_column = self.find_time_column(df)

                if time_column:
                    self.csv_info_label.config(
                        text=f"Loaded CSV: {os.path.basename(file_path)} (rows: {len(df)}, time col: {time_column})",
                        fg="green"
                    )
                    self.result_text.insert(tk.END, f"✅ CSV loaded successfully: {len(df)} rows, time column: '{time_column}'\n")

                    # Auto display CSV preview
                    try:
                        self.show_csv_preview(df, time_column)
                    except Exception as preview_error:
                        self.result_text.insert(tk.END, f"⚠️ Preview display error: {str(preview_error)}\n")
                else:
                    messagebox.showwarning("Warning", "Cannot find time information. Please check 'time' column or numeric data columns.")
                    self.existing_csv_path = None

            except Exception as e:
                error_details = str(e)
                if "encoding" in error_details.lower():
                    error_msg = f"CSV file encoding issue. Please use UTF-8 or ANSI encoded CSV file."
                elif "separator" in error_details.lower() or "delimiter" in error_details.lower():
                    error_msg = f"CSV file format issue. Please ensure it's comma-separated CSV file."
                else:
                    error_msg = f"CSV file processing error: {error_details}"

                messagebox.showerror("Error", error_msg)
                self.existing_csv_path = None

    def find_time_column(self, df):
        """Improved automatic time column detection - prioritize columns with changing values"""

        # 1. Check 'time' column first (exact name)
        if 'time' in df.columns:
            col = 'time'
            if pd.api.types.is_numeric_dtype(df[col]) and df[col].nunique() > 1:
                return col

        # 2. Find 'time' related column names (changing values only)
        for col in df.columns:
            if 'time' in col.lower():
                if pd.api.types.is_numeric_dtype(df[col]) and df[col].nunique() > 1:
                    # Check if it has time-like increasing pattern
                    if self.is_time_like_column(df[col]):
                        return col

        # 3. Check column f (6th column) - original logic
        if len(df.columns) > 5:
            col = df.columns[5]
            if pd.api.types.is_numeric_dtype(df[col]) and df[col].nunique() > 1:
                if self.is_time_like_column(df[col]):
                    return col

        # 4. Find numeric data columns with changing values
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]) and df[col].nunique() > 1:
                # Check if data looks like time (non-negative and increasing pattern)
                if self.is_time_like_column(df[col]):
                    return col

        return None

    def is_time_like_column(self, series):
        """Helper function to determine if column is time-like"""
        try:
            # Basic condition: numeric and values change
            if not pd.api.types.is_numeric_dtype(series) or series.nunique() <= 1:
                return False

            # If all values are same, not a time column
            if series.min() == series.max():
                return False

            # Values should be non-negative
            if series.min() < 0:
                return False

            # Check if mostly increasing pattern (time characteristic)
            diff = series.diff().dropna()
            if len(diff) > 0:
                positive_diff_ratio = (diff >= 0).sum() / len(diff)
                return positive_diff_ratio > 0.8  # 80% or more increasing

            return True

        except:
            return False

    def show_csv_preview(self, df, time_column):
        """Display CSV preview"""
        try:
            preview = f"\n📊 CSV Preview:\n"
            preview += f"Columns: {list(df.columns)}\n"

            # Safe time range calculation
            try:
                time_min = df[time_column].min()
                time_max = df[time_column].max()
                preview += f"Time range: {time_min:.3f}s ~ {time_max:.3f}s\n"
            except:
                preview += f"Time range: calculation failed\n"

            preview += f"Sample data (first 3 rows):\n"

            for i in range(min(3, len(df))):
                try:
                    time_val = df[time_column].iloc[i]
                    preview += f"  Row{i + 1}: time={time_val:.3f}s\n"
                except:
                    preview += f"  Row{i + 1}: time=read error\n"

            preview += "-" * 40 + "\n"
            self.result_text.insert(tk.END, preview)

        except Exception as e:
            error_msg = f"📊 CSV preview error: {str(e)}\n"
            self.result_text.insert(tk.END, error_msg)

    def add_shot_times_to_csv(self):
        """Add shot times to existing CSV - usability improvements"""
        if not self.existing_csv_path:
            messagebox.showwarning("Warning", "Please load a CSV file first.")
            return

        if not self.shot_data:
            messagebox.showwarning("Warning", "No shot analysis data. Please run video analysis first.")
            return

        try:
            # Load existing CSV - prevent encoding errors
            df = None
            encodings = ['utf-8-sig', 'utf-8', 'cp949', 'euc-kr', 'latin1']

            for encoding in encodings:
                try:
                    df = pd.read_csv(self.existing_csv_path, encoding=encoding)
                    break
                except:
                    continue

            if df is None:
                messagebox.showerror("Error", "Cannot read CSV file.")
                return

            time_column = self.find_time_column(df)

            if not time_column:
                messagebox.showerror("Error", "Cannot find time column.")
                return

            # Add shot time column (reset if already exists)
            df['shot_time'] = ''

            # Match with shot data
            tolerance = self.time_tolerance.get()
            matched_count = 0
            shot_details = []

            for i, shot in enumerate(self.shot_data):
                shot_time = shot['time_seconds']

                # Find rows within time tolerance
                time_diff = abs(df[time_column] - shot_time)
                matching_rows = time_diff <= tolerance

                if matching_rows.any():
                    df.loc[matching_rows, 'shot_time'] = 'O'
                    matched_rows_count = len(df[matching_rows])
                    matched_count += matched_rows_count
                    shot_details.append(f"Shot{i + 1}: {shot_time:.3f}s → {matched_rows_count} rows matched")

            # Save results
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = os.path.splitext(os.path.basename(self.existing_csv_path))[0]
            default_name = f"{base_name}_with_shots_{timestamp}.csv"

            save_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Save CSV with shot times"
            )

            if save_path:
                # Add default name if filename is empty
                if not save_path.endswith('.csv'):
                    save_path = save_path + '.csv'
                if os.path.basename(save_path) == '.csv':  # If filename is empty
                    dir_path = os.path.dirname(save_path)
                    save_path = os.path.join(dir_path, default_name)
                df.to_csv(save_path, index=False, encoding='utf-8-sig')

                # Detailed result message
                summary = f"\n{'=' * 60}\n"
                summary += f"📊 Shot times added successfully!\n"
                summary += f"📁 Original CSV: {os.path.basename(self.existing_csv_path)} ({len(df)} rows)\n"
                summary += f"🎯 Detected shots: {len(self.shot_data)}\n"
                summary += f"✅ Matched rows: {matched_count}\n"
                summary += f"⏰ Time tolerance: ±{tolerance}s\n"
                summary += f"💾 Save path: {save_path}\n\n"
                summary += "📋 Matching details:\n"
                for detail in shot_details[:5]:  # Show first 5 only
                    summary += f"  {detail}\n"
                if len(shot_details) > 5:
                    summary += f"  ... and {len(shot_details) - 5} more\n"
                summary += f"{'=' * 60}\n"

                self.result_text.insert(tk.END, summary)
                self.result_text.see(tk.END)

                messagebox.showinfo("Complete",
                                    f"CSV with shot times saved successfully!\n\n"
                                    f"Detected shots: {len(self.shot_data)}\n"
                                    f"Matched rows: {matched_count}\n"
                                    f"Match rate: {(matched_count / len(self.shot_data) * 100):.1f}%\n"
                                    f"Saved to: {os.path.basename(save_path)}")

        except Exception as e:
            error_details = str(e)
            if "initialvalue" in error_details:
                error_msg = "File save dialog error. Please enter filename manually."
            elif "encoding" in error_details.lower():
                error_msg = f"CSV file encoding issue. Try saving in different format.\nDetails: {error_details}"
            elif "permission" in error_details.lower():
                error_msg = "File permission error. Try saving to different location or check if file is open."
            else:
                error_msg = f"CSV processing error:\n{error_details}"

            messagebox.showerror("Error", error_msg)
            self.result_text.insert(tk.END, f"❌ Error: {error_msg}\n")

    def load_video(self):
        """Load video and extract sample frames"""
        file_path = filedialog.askopenfilename(
            title="Select video file",
            filetypes=[("Video files", "*.mp4 *.avi *.mov *.mkv *.wmv")]
        )
        if file_path:
            self.video_path = file_path
            self.extract_sample_frames()

    def extract_sample_frames(self):
        """Extract 10 random frames"""
        try:
            cap = cv2.VideoCapture(self.video_path)
            total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
            self.video_fps = cap.get(cv2.CAP_PROP_FPS)
            duration = total_frames / self.video_fps

            # Select 10 random frames (excluding first and last 10%)
            start_frame = int(total_frames * 0.1)
            end_frame = int(total_frames * 0.9)

            if end_frame - start_frame < 10:
                # If video too short, sample from entire video
                frame_indices = random.sample(range(0, total_frames), min(10, total_frames))
            else:
                frame_indices = random.sample(range(start_frame, end_frame), 10)

            frame_indices.sort()

            self.sample_frames = []
            for i, frame_idx in enumerate(frame_indices):
                cap.set(cv2.CAP_PROP_POS_FRAMES, frame_idx)
                ret, frame = cap.read()
                if ret:
                    frame_time = frame_idx / self.video_fps
                    self.sample_frames.append({
                        'index': i + 1,
                        'frame_number': frame_idx,
                        'time': frame_time,
                        'frame': frame.copy()
                    })

            cap.release()

            # Set first frame as default
            if self.sample_frames:
                self.current_frame = self.sample_frames[0]['frame']
                self.display_frame(self.current_frame)
                self.frame_info_label.config(
                    text=f"Selected Frame: #1 (frame:{self.sample_frames[0]['frame_number']}, {self.sample_frames[0]['time']:.1f}s)",
                    fg="green"
                )

            self.info_label.config(text=f"FPS: {self.video_fps:.1f} | Frames: {total_frames} | Duration: {duration:.1f}s")
            self.status_label.config(text=f"Status: Video loaded ({len(self.sample_frames)} sample frames ready)")

            self.result_text.insert(tk.END, f"🎬 Video loaded: {os.path.basename(self.video_path)}\n")
            self.result_text.insert(tk.END, f"📷 {len(self.sample_frames)} sample frames extracted\n")

        except Exception as e:
            messagebox.showerror("Error", f"Video loading error: {str(e)}")

    def select_frame(self):
        """Frame selection window"""
        if not self.sample_frames:
            messagebox.showwarning("Warning", "Please load a video first.")
            return

        # Create frame selection window
        frame_window = tk.Toplevel(self.root)
        frame_window.title("Frame Selection - Choose frame for region setup")
        frame_window.geometry("1200x800")

        # Frame display area
        canvas_frame = tk.Frame(frame_window)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Scrollable canvas
        canvas = tk.Canvas(canvas_frame)
        scrollbar = tk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Create frame buttons (2x5 grid)
        selected_frame = [None]  # Store selected frame

        for i, frame_data in enumerate(self.sample_frames):
            row = i // 5
            col = i % 5

            # Prepare frame image
            frame = frame_data['frame']
            height, width = frame.shape[:2]
            scale = min(200 / width, 150 / height)
            new_width = int(width * scale)
            new_height = int(height * scale)

            resized_frame = cv2.resize(frame, (new_width, new_height))
            rgb_frame = cv2.cvtColor(resized_frame, cv2.COLOR_BGR2RGB)
            pil_image = Image.fromarray(rgb_frame)
            photo = ImageTk.PhotoImage(pil_image)

            # Frame info
            frame_info = f"Frame {frame_data['index']}\nFrame#: {frame_data['frame_number']}\nTime: {frame_data['time']:.1f}s"

            # Frame button
            frame_btn = tk.Button(
                scrollable_frame,
                image=photo,
                text=frame_info,
                compound=tk.TOP,
                command=lambda fd=frame_data: self.on_frame_selected(fd, selected_frame, frame_window),
                relief=tk.RAISED,
                bd=2
            )
            frame_btn.image = photo  # Keep reference
            frame_btn.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")

        # Grid settings
        for i in range(5):
            scrollable_frame.columnconfigure(i, weight=1)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Bottom buttons
        btn_frame = tk.Frame(frame_window)
        btn_frame.pack(fill=tk.X, pady=10)

        tk.Button(btn_frame, text="❌ Cancel", command=frame_window.destroy,
                  bg="lightcoral", font=("Arial", 12)).pack(side=tk.RIGHT, padx=10)

        # Guide text
        info_label = tk.Label(frame_window,
                              text="Select a frame for region setup. Choose a frame where ammo information is clearly visible.",
                              font=("Arial", 11), fg="blue")
        info_label.pack(pady=5)

    def on_frame_selected(self, frame_data, selected_frame, window):
        """Called when frame is selected"""
        selected_frame[0] = frame_data
        self.current_frame = frame_data['frame']
        self.display_frame(self.current_frame)

        self.frame_info_label.config(
            text=f"Selected Frame: #{frame_data['index']} (frame:{frame_data['frame_number']}, {frame_data['time']:.1f}s)",
            fg="green"
        )

        self.result_text.insert(tk.END, f"📷 Frame #{frame_data['index']} selected ({frame_data['time']:.1f}s)\n")

        window.destroy()
        messagebox.showinfo("Complete", f"Frame #{frame_data['index']} selected.\nNow proceed with region setup.")

    def display_frame(self, frame):
        """Display frame in GUI"""
        height, width = frame.shape[:2]
        max_width, max_height = 900, 450

        scale = min(max_width / width, max_height / height)
        new_width = int(width * scale)
        new_height = int(height * scale)

        resized_frame = cv2.resize(frame, (new_width, new_height))
        rgb_frame = cv2.cvtColor(resized_frame, cv2.COLOR_BGR2RGB)
        pil_image = Image.fromarray(rgb_frame)
        photo = ImageTk.PhotoImage(pil_image)

        self.video_label.config(image=photo, text="")
        self.video_label.image = photo

    def set_total_ammo_region(self):
        self.total_ammo_region = self.select_region("Select total ammo region")
        if self.total_ammo_region:
            self.result_text.insert(tk.END, f"✅ Total ammo region set: {self.total_ammo_region}\n")

    def set_current_ammo_region(self):
        self.current_ammo_region = self.select_region("Select current ammo region")
        if self.current_ammo_region:
            self.result_text.insert(tk.END, f"✅ Current ammo region set: {self.current_ammo_region}\n")

    def select_region(self, title):
        """Region selection - greatly improved usability"""
        if self.current_frame is None:
            messagebox.showwarning("Warning", "Please load video and select frame first.")
            return None

        try:
            frame = self.current_frame.copy()
        except Exception as e:
            messagebox.showerror("Error", f"Frame loading error: {e}")
            return None

        # New window for region selection - larger window
        region_window = tk.Toplevel(self.root)
        region_window.title(f"{title}")
        region_window.geometry("1400x900")
        region_window.resizable(True, True)

        # Main frame
        main_frame = tk.Frame(region_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Top control frame
        control_frame = tk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, 10))

        # Zoom controls
        zoom_frame = tk.Frame(control_frame)
        zoom_frame.pack(side=tk.LEFT)

        tk.Label(zoom_frame, text="Zoom:", font=("Arial", 10, "bold")).pack(side=tk.LEFT)

        zoom_scale = tk.DoubleVar(value=1.0)
        zoom_spinbox = tk.Spinbox(zoom_frame, from_=0.5, to=3.0, increment=0.1,
                                  textvariable=zoom_scale, width=6)
        zoom_spinbox.pack(side=tk.LEFT, padx=5)

        tk.Button(zoom_frame, text="🔍 Apply",
                  command=lambda: self.update_canvas_zoom(canvas, frame, zoom_scale.get(), photo_holder, selection),
                  bg="lightblue").pack(side=tk.LEFT, padx=5)

        # Help
        help_frame = tk.Frame(control_frame)
        help_frame.pack(side=tk.RIGHT)

        help_text = "💡 Left-click drag: Select region | Right-click drag: Move | Wheel: Zoom"
        tk.Label(help_frame, text=help_text, font=("Arial", 9), fg="blue").pack()

        # Status display
        status_frame = tk.Frame(control_frame)
        status_frame.pack()

        drag_status = tk.Label(status_frame, text="Status: Ready", font=("Arial", 9), fg="darkgreen")
        drag_status.pack()

        # Canvas frame (scrollable)
        canvas_frame = tk.Frame(main_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        # Scrollbars
        h_scrollbar = tk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL)
        v_scrollbar = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL)

        # Canvas
        canvas = tk.Canvas(canvas_frame,
                           xscrollcommand=h_scrollbar.set,
                           yscrollcommand=v_scrollbar.set,
                           bg="white")

        h_scrollbar.config(command=canvas.xview)
        v_scrollbar.config(command=canvas.yview)

        # Pack
        canvas.grid(row=0, column=0, sticky="nsew")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")

        canvas_frame.grid_rowconfigure(0, weight=1)
        canvas_frame.grid_columnconfigure(0, weight=1)

        # Display image
        image_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        pil_image = Image.fromarray(image_rgb)
        original_size = pil_image.size

        # Initial scaling (fit to window)
        canvas_width = 1200
        canvas_height = 600
        scale = min(canvas_width / original_size[0], canvas_height / original_size[1])
        initial_size = (int(original_size[0] * scale), int(original_size[1] * scale))

        pil_image = pil_image.resize(initial_size, Image.Resampling.LANCZOS)
        photo_holder = [ImageTk.PhotoImage(pil_image)]  # List to keep reference

        image_item = canvas.create_image(0, 0, anchor=tk.NW, image=photo_holder[0])
        canvas.configure(scrollregion=canvas.bbox("all"))

        # Selection variables
        selection = {
            'start_x': 0, 'start_y': 0, 'end_x': 0, 'end_y': 0,
            'rect_id': None, 'confirmed': False, 'dragging': False,
            'current_scale': scale, 'pan_mode': False
        }

        # Mouse event handlers (fixed)
        def on_left_click(event):
            canvas_x = canvas.canvasx(event.x)
            canvas_y = canvas.canvasy(event.y)
            selection['start_x'], selection['start_y'] = canvas_x, canvas_y
            selection['dragging'] = True
            drag_status.config(text="Status: Selecting region...", fg="red")
            if selection['rect_id']:
                canvas.delete(selection['rect_id'])

        def on_left_drag(event):
            if selection['dragging']:
                canvas_x = canvas.canvasx(event.x)
                canvas_y = canvas.canvasy(event.y)
                selection['end_x'], selection['end_y'] = canvas_x, canvas_y

                if selection['rect_id']:
                    canvas.delete(selection['rect_id'])

                selection['rect_id'] = canvas.create_rectangle(
                    selection['start_x'], selection['start_y'],
                    selection['end_x'], selection['end_y'],
                    outline="red", width=3, tags="selection"
                )

                # Real-time size display
                width = abs(selection['end_x'] - selection['start_x'])
                height = abs(selection['end_y'] - selection['start_y'])
                drag_status.config(text=f"Dragging: {width:.0f} x {height:.0f} pixels", fg="blue")

        def on_left_release(event):
            selection['dragging'] = False
            if selection['rect_id']:
                width = abs(selection['end_x'] - selection['start_x'])
                height = abs(selection['end_y'] - selection['start_y'])
                drag_status.config(text=f"Region selected: {width:.0f} x {height:.0f} pixels", fg="green")
            else:
                drag_status.config(text="Status: Ready", fg="darkgreen")

        def on_right_click(event):
            canvas.scan_mark(event.x, event.y)
            selection['pan_mode'] = True
            drag_status.config(text="Status: Moving image...", fg="orange")

        def on_right_drag(event):
            if selection['pan_mode']:
                canvas.scan_dragto(event.x, event.y, gain=1)

        def on_right_release(event):
            selection['pan_mode'] = False
            drag_status.config(text="Status: Ready", fg="darkgreen")

        def on_mouse_wheel(event):
            # Mouse wheel zoom
            try:
                if event.delta > 0 or event.num == 4:  # Scroll up
                    factor = 1.1
                elif event.delta < 0 or event.num == 5:  # Scroll down
                    factor = 0.9
                else:
                    return

                new_zoom = zoom_scale.get() * factor
                new_zoom = max(0.1, min(5.0, new_zoom))
                zoom_scale.set(new_zoom)
                self.update_canvas_zoom(canvas, frame, new_zoom, photo_holder, selection)
            except:
                pass

        # Event binding (separated)
        canvas.bind("<Button-1>", on_left_click)
        canvas.bind("<B1-Motion>", on_left_drag)
        canvas.bind("<ButtonRelease-1>", on_left_release)

        canvas.bind("<Button-3>", on_right_click)
        canvas.bind("<B3-Motion>", on_right_drag)
        canvas.bind("<ButtonRelease-3>", on_right_release)

        canvas.bind("<MouseWheel>", on_mouse_wheel)
        # Linux mouse wheel events
        canvas.bind("<Button-4>", on_mouse_wheel)
        canvas.bind("<Button-5>", on_mouse_wheel)

        canvas.focus_set()  # Keyboard focus

        # Bottom button frame
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        def confirm_selection():
            if selection['rect_id'] and (abs(selection['end_x'] - selection['start_x']) > 10 and
                                         abs(selection['end_y'] - selection['start_y']) > 10):
                selection['confirmed'] = True
                drag_status.config(text="Status: Selection confirmed!", fg="green")
                region_window.destroy()
            else:
                messagebox.showwarning("Warning", "Please select a sufficiently large region.\n(minimum 10x10 pixels)")
                drag_status.config(text="Status: Region too small", fg="red")

        def reset_selection():
            if selection['rect_id']:
                canvas.delete(selection['rect_id'])
                selection['rect_id'] = None
                drag_status.config(text="Status: Selection reset", fg="darkgreen")

        # Buttons
        tk.Button(button_frame, text="🔄 Reset Selection", command=reset_selection,
                  bg="lightyellow", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)

        # Test dummy region button
        def create_test_rect():
            if selection['rect_id']:
                canvas.delete(selection['rect_id'])
            selection['start_x'], selection['start_y'] = 50, 50
            selection['end_x'], selection['end_y'] = 200, 100
            selection['rect_id'] = canvas.create_rectangle(
                selection['start_x'], selection['start_y'],
                selection['end_x'], selection['end_y'],
                outline="red", width=3, tags="selection"
            )
            drag_status.config(text="Test region created: 150 x 50 pixels", fg="green")

        tk.Button(button_frame, text="🧪 Test Region", command=create_test_rect,
                  bg="lightgray", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)

        tk.Button(button_frame, text="❌ Cancel", command=region_window.destroy,
                  bg="lightcoral", font=("Arial", 12)).pack(side=tk.RIGHT, padx=5)

        tk.Button(button_frame, text="✅ Confirm Selection", command=confirm_selection,
                  bg="lightgreen", font=("Arial", 12, "bold")).pack(side=tk.RIGHT, padx=5)

        # Guide text
        instruction_text = "💡 Left-click drag: Select region | Right-click drag: Move image | Try 🧪 Test Region first"
        instruction_label = tk.Label(region_window, text=instruction_text,
                                     font=("Arial", 9), fg="blue")
        instruction_label.pack(pady=5)

        # Selection info display
        info_label = tk.Label(button_frame, text="Please select a region",
                              font=("Arial", 10), fg="darkblue")
        info_label.pack(side=tk.LEFT, padx=20)

        region_window.wait_window()

        if selection['confirmed']:
            # Convert to actual image coordinates
            current_scale = selection['current_scale']

            x1 = int(min(selection['start_x'], selection['end_x']) / current_scale)
            y1 = int(min(selection['start_y'], selection['end_y']) / current_scale)
            x2 = int(max(selection['start_x'], selection['end_x']) / current_scale)
            y2 = int(max(selection['start_y'], selection['end_y']) / current_scale)

            # Boundary check
            x1 = max(0, min(x1, original_size[0]))
            y1 = max(0, min(y1, original_size[1]))
            x2 = max(0, min(x2, original_size[0]))
            y2 = max(0, min(y2, original_size[1]))

            return (x1, y1, x2, y2)
        return None

    def update_canvas_zoom(self, canvas, frame, zoom_factor, photo_holder, selection):
        """Update canvas zoom"""
        try:
            image_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            pil_image = Image.fromarray(image_rgb)
            original_size = pil_image.size

            new_size = (int(original_size[0] * zoom_factor), int(original_size[1] * zoom_factor))
            pil_image = pil_image.resize(new_size, Image.Resampling.LANCZOS)

            photo_holder[0] = ImageTk.PhotoImage(pil_image)

            # Update all image items in canvas
            canvas.delete("image")
            canvas.create_image(0, 0, anchor=tk.NW, image=photo_holder[0], tags="image")
            canvas.configure(scrollregion=canvas.bbox("all"))

            # Update current scale
            selection['current_scale'] = zoom_factor

        except Exception as e:
            print(f"Zoom update error: {e}")

    def play_alert_sound(self):
        """Play alert sound when shot detected"""
        try:
            import winsound
            winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
        except:
            try:
                import os
                os.system("echo \a")
            except:
                pass

    def extract_number_from_region(self, frame, region):
        """Extract number using OCR - improved version"""
        if region is None:
            return None

        x1, y1, x2, y2 = region
        roi = frame[y1:y2, x1:x2]

        if roi.size == 0:
            return None

        try:
            # Enhanced image preprocessing
            gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)

            # Enlarge for better OCR accuracy
            if roi.shape[0] < 50 or roi.shape[1] < 50:
                gray = cv2.resize(gray, None, fx=4, fy=4, interpolation=cv2.INTER_CUBIC)

            # Noise removal and binarization
            gray = cv2.medianBlur(gray, 3)

            # Try multiple thresholds
            thresholds = [
                cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1],
                cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)[1],
                cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
            ]

            # OCR settings
            configs = [
                r'--oem 3 --psm 8 -c tessedit_char_whitelist=0123456789',
                r'--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789',
                r'--oem 3 --psm 13 -c tessedit_char_whitelist=0123456789'
            ]

            for thresh in thresholds:
                # Morphological operations
                kernel = np.ones((2, 2), np.uint8)
                thresh = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)
                thresh = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel)

                for config in configs:
                    try:
                        text = pytesseract.image_to_string(thresh, config=config)
                        numbers = re.findall(r'\d+', text.strip())
                        if numbers:
                            return int(numbers[0])
                    except:
                        continue

        except Exception as e:
            pass

        return None

    def analyze_video(self):
        """Video analysis - keep existing method"""
        if not self.current_ammo_region:
            messagebox.showwarning("Warning", "Please set current ammo region first.")
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

            if frame_count % skip_frames != 0:
                continue

            self.progress_bar['value'] = frame_count

            current_ammo = self.extract_number_from_region(frame, self.current_ammo_region)
            total_ammo = self.extract_number_from_region(frame, self.total_ammo_region)

            self.root.after(0, self.update_ammo_display, total_ammo, current_ammo, frame_count)

            # Shot detection
            if current_ammo is not None and previous_ammo is not None:
                if current_ammo < previous_ammo:
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

                    result = f"🎯 {len(self.shot_data):2d}. {shot_time} | {previous_ammo}→{current_ammo} ({previous_ammo - current_ammo} shots)\n"
                    self.result_text.insert(tk.END, result)
                    self.result_text.see(tk.END)

                    if self.sound_alert.get():
                        self.play_alert_sound()

                    self.root.after(0, self.update_shot_count, len(self.shot_data))
                    self.root.update()

            if current_ammo is not None:
                previous_ammo = current_ammo

            if frame_count % 100 == 0:
                elapsed = time.time() - start_time
                fps_actual = frame_count / elapsed if elapsed > 0 else 0
                self.status_label.config(
                    text=f"Analyzing... {frame_count}/{total_frames} ({fps_actual:.1f} FPS)"
                )
                self.root.update()

        cap.release()

        if self.analysis_running:
            elapsed = time.time() - start_time
            self.status_label.config(
                text=f"Analysis complete! {len(self.shot_data)} shots detected (time: {elapsed:.1f}s)"
            )

            summary = f"\n{'=' * 50}\n"
            summary += f"📊 Analysis complete: {len(self.shot_data)} shots detected\n"
            summary += f"⏱️ Analysis time: {elapsed:.1f}s\n"
            summary += f"🎬 Total frames: {total_frames} (FPS: {fps:.1f})\n"
            summary += f"{'=' * 50}\n"
            self.result_text.insert(tk.END, summary)

        self.analysis_running = False

    def update_ammo_display(self, total_ammo, current_ammo, frame_count):
        """Update real-time ammo info display"""
        total_str = str(total_ammo) if total_ammo is not None else "?"
        current_str = str(current_ammo) if current_ammo is not None else "?"

        self.current_ammo_label.config(text=f"Current Ammo: {current_str}")
        self.total_ammo_label.config(text=f"Total Ammo: {total_str}")

    def update_shot_count(self, count):
        """Update shot count"""
        self.shot_count_label.config(text=f"Detected Shots: {count}")

    def start_analysis(self):
        """Start analysis"""
        if not self.video_path:
            messagebox.showwarning("Warning", "Please select a video file first.")
            return

        if not self.current_ammo_region:
            messagebox.showwarning("Warning", "Please set current ammo region.")
            return

        self.analysis_running = True
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "🚀 Analysis started...\n\n")

        self.analysis_thread = threading.Thread(target=self.analyze_video, daemon=True)
        self.analysis_thread.start()

    def stop_analysis(self):
        """Stop analysis"""
        self.analysis_running = False
        self.status_label.config(text="Status: Analysis stopped")

    def run(self):
        """Run program"""
        self.root.mainloop()


if __name__ == "__main__":
    try:
        import cv2
        import pytesseract
        import pandas as pd
    except ImportError as e:
        print("Please install required libraries:")
        print("pip install opencv-python pytesseract pillow pandas")
        print("Also install Tesseract OCR:")
        print("https://github.com/tesseract-ocr/tesseract")
        exit(1)

    app = VideoAmmoAnalyzer()
    app.run()