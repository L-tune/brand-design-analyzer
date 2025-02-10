import os
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext, messagebox
from pdf2image import convert_from_path
import time
import threading
from PIL import Image as PILImage
import numpy as np
import json
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
import io
import logging
import datetime
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

def check_dependencies():
    required_packages = {
        'pdf2image': 'pdf2image',
        'openai': 'openai>=1.0.0',  # Указываем минимальную версию
        'PIL': 'Pillow',
        'numpy': 'numpy',
        'reportlab': 'reportlab'
    }
    
    missing_packages = []
    for module, package in required_packages.items():
        try:
            __import__(module)
            if module == 'openai':
                import pkg_resources
                version = pkg_resources.get_distribution('openai').version
                if pkg_resources.parse_version(version) < pkg_resources.parse_version('1.0.0'):
                    missing_packages.append(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        raise ImportError(f"Отсутствуют необходимые пакеты: {', '.join(missing_packages)}. "
                         f"Установите их с помощью pip install {' '.join(missing_packages)}")

class BrandAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Brand Design Analyzer Pro")
        self.root.geometry("1200x800")
        
        # Настраиваем логирование
        self.setup_logging()
        
        # Получаем API ключ из переменной окружения
        api_key = os.getenv('OPENAI_API_KEY')
        if not api_key:
            self.show_error("Не найден API ключ OpenAI. Установите переменную окружения OPENAI_API_KEY")
            raise ValueError("Missing OpenAI API key")
        
        # Инициализация OpenAI клиента
        try:
            self.client = OpenAI(
                api_key=api_key,
                default_headers={"OpenAI-Beta": "assistants=v1"}
            )
        except Exception as e:
            self.log_event(f"Ошибка инициализации OpenAI клиента: {str(e)}", level='error')
            raise
        
        # Создаем интерфейс
        self.create_widgets()
        
        # Инициализируем контексты
        self.analysis_context = {
            'design_systems': [],
            'variants': {},
            'elements': {},
            'slides_map': {}
        }
        
        self.presentation_context = {
            'key_elements': {},
            'design_decisions': [],
            'story_flow': [],
            'last_comments': []
        }
        
        self.start_time = 0
        
        # Логируем запуск после создания всех компонентов
        self.log_event("Приложение запущено")
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def create_widgets(self):
        # Верхняя панель с кнопками
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Кнопка выбора файла и поле с путём
        self.select_button = ttk.Button(
            top_frame, 
            text="Выбрать PDF файл", 
            command=self.select_file
        )
        self.select_button.pack(side=tk.LEFT, padx=5)
        
        self.file_path_var = tk.StringVar()
        self.file_path_entry = ttk.Entry(
            top_frame, 
            textvariable=self.file_path_var,
            width=50,
            state='readonly'
        )
        self.file_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Панель контекста
        context_frame = ttk.LabelFrame(self.root, text="Контекст анализа", padding="5")
        context_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Поле для ввода контекста
        self.context_text = scrolledtext.ScrolledText(
            context_frame, 
            height=4, 
            wrap=tk.WORD
        )
        self.context_text.pack(fill=tk.X, expand=True)
        
        # Добавляем подсказку
        self.default_context = """Опишите контекст анализа. Например:
Это презентация ребрендинга компании X, которая работает в сфере Y. 
Основная цель редизайна - Z. Целевая аудитория - A."""
        
        self.context_text.insert('1.0', self.default_context)
        
        # Биндим стандартные сочетания клавиш
        self.context_text.bind('<Control-a>', self.select_all)
        self.context_text.bind('<Control-A>', self.select_all)
        self.context_text.bind('<Control-v>', self.paste_text)
        self.context_text.bind('<Control-V>', self.paste_text)
        self.context_text.bind('<Control-c>', self.copy_text)
        self.context_text.bind('<Control-C>', self.copy_text)
        self.context_text.bind('<Control-x>', self.cut_text)
        self.context_text.bind('<Control-X>', self.cut_text)
        
        # Добавляем контекстное меню
        self.context_menu = tk.Menu(self.context_text, tearoff=0)
        self.context_menu.add_command(label="Вырезать", command=lambda: self.context_text.event_generate('<<Cut>>'))
        self.context_menu.add_command(label="Копировать", command=lambda: self.context_text.event_generate('<<Copy>>'))
        self.context_menu.add_command(label="Вставить", command=lambda: self.context_text.event_generate('<<Paste>>'))
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Выделить всё", command=lambda: self.context_text.event_generate('<<SelectAll>>'))
        
        # Биндим появление контекстного меню
        self.context_text.bind('<Button-3>', self.show_context_menu)
        
        # Остальные бинды
        self.context_text.bind('<FocusIn>', self.on_context_focus_in)
        self.context_text.bind('<FocusOut>', self.on_context_focus_out)
        
        # Кнопка анализа
        self.analyze_button = ttk.Button(
            context_frame,
            text="Анализировать презентацию",
            command=self.analyze_all
        )
        self.analyze_button.pack(pady=(5,0))
        
        # Добавляем кнопку тестового анализа
        self.test_button = ttk.Button(
            context_frame,
            text="Тестовый анализ (10 слайдов)",
            command=self.test_analyze_first_10
        )
        self.test_button.pack(pady=(5,0))
        
        # Добавляем прогресс-бар
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            context_frame,
            variable=self.progress_var,
            maximum=100
        )
        self.progress_bar.pack(fill=tk.X, pady=(5,0))
        
        # Текстовое поле для вывода
        self.result_text = scrolledtext.ScrolledText(self.root, wrap=tk.WORD)
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Статус-бар
        self.status_label = ttk.Label(
            self.root, 
            text="Готов к работе", 
            relief=tk.SUNKEN, 
            anchor=tk.W
        )
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=2)
        
    def select_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("PDF files", "*.pdf")]
        )
        if file_path:
            # Проверка размера файла (например, максимум 50 МБ)
            try:
                file_size = os.path.getsize(file_path) / (1024 * 1024)  # в МБ
                if file_size > 50:
                    self.show_error("Файл слишком большой. Максимальный размер: 50 МБ")
                    return
                self.file_path_var.set(file_path)
            except Exception as e:
                self.log_event(f"Ошибка при проверке файла: {str(e)}", level='error')
            
    def convert_pdf_to_images(self, pdf_path, output_folder):
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        
        try:
            images = convert_from_path(pdf_path)
            image_paths = []
            
            for i, image in enumerate(images):
                # Изменяем размер изображения в соответствии с рекомендациями API
                width, height = image.size
                # Максимальные размеры согласно документации
                max_size = (2000, 2000)
                
                # Масштабируем изображение, сохраняя пропорции
                if width > max_size[0] or height > max_size[1]:
                    image.thumbnail(max_size, Image.LANCZOS)
                
                image_path = os.path.join(output_folder, f'slide_{i+1}.jpg')
                image.save(image_path, 'JPEG', quality=95)
                
                # Проверяем размер файла (максимум 20MB согласно документации)
                if os.path.getsize(image_path) > 20 * 1024 * 1024:  # 20MB в байтах
                    self.log_event(f"Слайд {i+1} превышает максимальный размер 20MB, уменьшаем качество")
                    image.save(image_path, 'JPEG', quality=85)  # Уменьшаем качество
                
                image_paths.append(image_path)
            
            return image_paths
            
        except Exception as e:
            self.log_event(f"Ошибка при конвертации PDF: {str(e)}", level='error')
            raise
        
    def encode_image_to_base64(self, image_path):
        import base64
        with open(image_path, 'rb') as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')
            
    def initial_analysis(self, image_paths):
        """Первичный анализ всей презентации"""
        self.update_status("Проводим первичный анализ презентации...")
        
        system_prompt = """Вы - опытный арт-директор и бренд-аналитик. Проведите первичный анализ слайда и верните результат в JSON формате.
        Определите:
        1. Категорию слайда (концепция/элемент системы/применение/вариант)
        2. Если это вариант - к какой группе вариантов относится
        3. Ключевые визуальные элементы
        4. Связь с другими элементами системы"""
        
        initial_analysis = {}
        
        for i, image_path in enumerate(image_paths, 1):
            if self.is_text_slide(image_path):
                continue
                
            try:
                base64_image = self.encode_image_to_base64(image_path)
                response = self.client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[
                        {
                            "role": "system",
                            "content": system_prompt
                        },
                        {
                            "role": "user",
                            "content": [
                                {
                                    "type": "text",
                                    "text": "Проанализируйте этот слайд и предоставьте результат в JSON."
                                },
                                {
                                    "type": "image_url",
                                    "image_url": {
                                        "url": f"data:image/jpeg;base64,{base64_image}",
                                        "detail": "high"
                                    }
                                }
                            ]
                        }
                    ],
                    max_tokens=500,
                    response_format={ "type": "json_object" }
                )
                
                analysis = json.loads(response.choices[0].message.content)
                initial_analysis[i] = analysis
                self.update_status(f"Проанализирован слайд {i}")
                
            except Exception as e:
                self.log_event(f"Ошибка при первичном анализе слайда {i}: {str(e)}")
                
        return initial_analysis

    def build_smart_context(self, initial_analysis):
        """Создание умного контекста на основе первичного анализа"""
        self.update_status("Формируем общее понимание дизайн-системы...")
        
        context_prompt = """Проанализируйте предоставленные данные и создайте JSON-отчет со следующей структурой:
        1. Основные дизайн-системы
        2. Группы вариантов дизайна
        3. Ключевые принципы и паттерны
        4. Связи между элементами

        Данные для анализа: {analysis}"""
        
        try:
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {
                        "role": "system",
                        "content": context_prompt.format(analysis=json.dumps(initial_analysis, ensure_ascii=False))
                    },
                    {
                        "role": "user",
                        "content": "Создайте структурированный JSON-отчет на основе этих данных."
                    }
                ],
                max_tokens=1000,
                response_format={ "type": "json_object" }
            )
            
            return json.loads(response.choices[0].message.content)
            
        except Exception as e:
            self.log_event(f"Ошибка при создании умного контекста: {str(e)}")
            return {}

    def analyze_slide_with_context(self, image_path, slide_number, smart_context):
        max_retries = 3
        retry_delay = 2
        
        for attempt in range(max_retries):
            try:
                if self.is_text_slide(image_path):
                    self.log_event(f"Слайд {slide_number} пропущен (текстовый)")
                    return None
                
                base64_image = self.encode_image_to_base64(image_path)
                
                context = self.context_text.get('1.0', tk.END).strip()
                if context == self.default_context.strip():
                    context = "Анализ дизайна презентации"
                
                previous_context = self.get_brief_context()
                
                messages = [
                    {
                        "role": "system",
                        "content": f"""Вы - опытный арт-директор, представляющий концепцию дизайна клиенту в неформальной обстановке. 
                        Контекст проекта: {context}
                        Что мы уже обсудили: {previous_context}"""
                    },
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": "Расскажите про это решение просто и по делу."
                            },
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/jpeg;base64,{base64_image}",
                                    "detail": "high"
                                }
                            }
                        ]
                    }
                ]
                
                response = self.client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=messages,
                    max_tokens=500,
                    timeout=30  # Добавляем тайм-аут
                )
                
                analysis = response.choices[0].message.content
                
                self.update_presentation_context(slide_number, analysis)
                self.log_event(f"Слайд {slide_number} успешно проанализирован")
                return analysis
                
            except Exception as e:
                self.log_event(f"Попытка {attempt + 1}/{max_retries} для слайда {slide_number} не удалась: {str(e)}", level='error')
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    raise  # Пробрасываем ошибку после всех попыток

    def get_brief_context(self):
        """Формирует краткий контекст из предыдущих слайдов"""
        if not self.presentation_context['last_comments']:
            return "Это первый слайд презентации."
            
        context_parts = []
        
        # Добавляем последние комментарии
        if self.presentation_context['last_comments']:
            last_comments = self.presentation_context['last_comments'][-2:]  # Берем последние 2 комментария
            context_parts.append("Последние обсуждения:")
            for i, comment in enumerate(last_comments, 1):
                context_parts.append(f"- {comment}")
        
        # Добавляем ключевые элементы, если они есть
        if self.presentation_context['key_elements']:
            context_parts.append("\nКлючевые элементы дизайна:")
            for element, description in self.presentation_context['key_elements'].items():
                context_parts.append(f"- {element}: {description}")
        
        return "\n".join(context_parts)

    def update_presentation_context(self, slide_number, analysis):
        """Обновляет контекст презентации на основе нового анализа"""
        # Сохраняем комментарий
        self.presentation_context['last_comments'].append(analysis)
        
        try:
            # Анализируем комментарий для извлечения ключевой информации
            context_update_prompt = f"""
            Проанализируйте этот комментарий к слайду {slide_number} и верните строго валидный JSON с такой структурой:
            {{
                "key_elements": {{"element_name": "description"}},
                "design_decisions": ["decision1", "decision2"],
                "connections": ["connection1", "connection2"]
            }}

            Комментарий для анализа:
            {analysis}
            """
            
            response = self.client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {
                        "role": "system",
                        "content": "Вы - парсер, который создает только валидный JSON. Всегда проверяйте закрытие кавычек и скобок."
                    },
                    {
                        "role": "user",
                        "content": context_update_prompt
                    }
                ],
                max_tokens=500,
                response_format={ "type": "json_object" }
            )
            
            try:
                update_info = json.loads(response.choices[0].message.content)
                
                # Обновляем ключевые элементы
                if 'key_elements' in update_info:
                    self.presentation_context['key_elements'].update(update_info['key_elements'])
                
                # Добавляем дизайнерские решения
                if 'design_decisions' in update_info:
                    self.presentation_context['design_decisions'].extend(update_info['design_decisions'])
                
                # Обновляем поток повествования
                self.presentation_context['story_flow'].append({
                    'slide': slide_number,
                    'summary': analysis[:100] + '...' if len(analysis) > 100 else analysis
                })
                
            except json.JSONDecodeError as e:
                self.log_event(f"Ошибка парсинга JSON при обновлении контекста: {str(e)}", level='warning')
                
        except Exception as e:
            self.log_event(f"Ошибка при обновлении контекста: {str(e)}", level='error')

    def is_text_slide(self, image_path):
        try:
            with PILImage.open(image_path) as img:
                img = img.convert('L')
                img_array = np.array(img)
                text_pixels = np.sum(img_array < 128)
                total_pixels = img_array.size
                text_ratio = text_pixels / total_pixels
                return text_ratio > 0.15
        except Exception as e:
            self.log_event(f"Ошибка при анализе текстового слайда: {str(e)}", level='error')
            return False

    def analyze_all(self):
        """Основной метод анализа"""
        if not self.file_path_var.get():
            self.show_error("Выберите PDF файл для анализа")
            return
        
        self.result_text.delete('1.0', tk.END)
        self.result_text.insert(tk.END, "Начинаем комплексный анализ презентации...\n\n")
        self.result_text.update()
        
        thread = threading.Thread(target=self._analyze_all_slides)
        thread.start()

    def save_analysis_report(self, content, pdf_path):
        """Сохраняет отчет рядом с исходным PDF"""
        try:
            # Получаем путь и имя исходного файла
            pdf_dir = os.path.dirname(pdf_path)
            pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
            
            # Формируем имя файла отчета
            timestamp = time.strftime("%Y%m%d-%H%M%S")
            report_name = f"{pdf_name}_analysis_{timestamp}.txt"
            report_path = os.path.join(pdf_dir, report_name)
            
            # Добавляем информацию о контексте анализа
            context = self.context_text.get('1.0', tk.END).strip()
            if context == self.default_context.strip():
                context = "Стандартный анализ дизайна презентации"
            
            header = f"""АНАЛИЗ ПРЕЗЕНТАЦИИ: {pdf_name}
Дата анализа: {time.strftime("%Y-%m-%d %H:%M:%S")}

КОНТЕКСТ АНАЛИЗА:
{context}

💡 ПОДСКАЗКА ДЛЯ ПРЕЗЕНТАЦИИ:
- Текст структурирован для удобного чтения
- **Жирным** выделены ключевые слова
- Каждый слайд содержит главную мысль и детали
- Используйте • ГЛАВНОЕ как опорную точку
- В • ДЕТАЛИ собраны аргументы для обсуждения
- • СВЯЗЬ поможет создать плавный переход

РЕЗУЛЬТАТЫ АНАЛИЗА:
=================

"""
            
            # Сохраняем отчет
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(header + content)
            
            self.log_event(f"\nОтчет сохранен: {report_path}")
            return report_path
            
        except Exception as e:
            self.log_event(f"\nОшибка при сохранении отчета: {str(e)}")
            return None

    def _analyze_all_slides(self):
        try:
            pdf_path = self.file_path_var.get()
            self.log_event(f"Анализируемый файл: {pdf_path}")
            
            # Конвертируем PDF в изображения
            self.update_status("Конвертируем PDF в изображения...")
            output_folder = "slides_images"
            image_paths = self.convert_pdf_to_images(pdf_path, output_folder)
            total_slides = len(image_paths)
            
            self.log_event(f"Создано {total_slides} изображений слайдов")
            
            analysis_results = []  # Сохраняем результаты анализа
            
            # Анализируем каждый слайд
            for i, image_path in enumerate(image_paths, 1):
                self.log_event(f"Начинаем анализ слайда {i}")
                try:
                    analysis = self.analyze_slide_with_context(image_path, i, None)
                    if analysis:
                        self.update_interface(f"• Слайд {i}: {analysis}")
                        analysis_results.append((i, analysis))
                except Exception as e:
                    self.update_interface(f"• Слайд {i}: Ошибка при анализе слайда {i}: {str(e)}")
                
                # Обновляем прогресс
                progress = (i / total_slides) * 100
                self.progress_var.set(progress)
                self.update_status(f"Проанализировано {i} из {total_slides} слайдов ({progress:.1f}%)")
            
            # Создаем итоговый отчет
            if analysis_results:
                self.update_status("Создаем итоговый отчет...")
                
                # Сохраняем текстовый отчет
                report_content = "\n".join([f"Слайд {i}: {analysis}" for i, analysis in analysis_results])
                report_path = self.save_analysis_report(report_content, pdf_path)
                
                # Создаем презентационный гайд
                try:
                    guide_path = self.create_presentation_guide(report_content, pdf_path, image_paths)
                    self.log_event(f"\nПрезентационный гайд сохранен: {guide_path}")
                except Exception as e:
                    self.log_event(f"\nОшибка при создании презентационного гайда: {str(e)}", level='error')
            
            self.update_status("Анализ завершен")
            
        except Exception as e:
            self.log_event(f"Ошибка при анализе: {str(e)}", level='error')
            self.show_error(f"Произошла ошибка при анализе: {str(e)}")
        finally:
            # Очищаем временные файлы
            try:
                for image_path in image_paths:
                    if os.path.exists(image_path):
                        os.remove(image_path)
                if os.path.exists(output_folder):
                    os.rmdir(output_folder)
            except Exception as e:
                self.log_event(f"Ошибка при очистке временных файлов: {str(e)}", level='warning')

    def show_error(self, message):
        messagebox.showerror("Ошибка", message)

    def estimate_time_left(self, current, total):
        if current == 0:
            return "..."
        
        elapsed_time = time.time() - self.start_time
        time_per_slide = elapsed_time / current
        remaining_slides = total - current
        remaining_time = remaining_slides * time_per_slide
        
        if remaining_time < 60:
            return f"{int(remaining_time)} сек"
        else:
            return f"{int(remaining_time / 60)} мин {int(remaining_time % 60)} сек"

    def on_context_focus_in(self, event):
        """Очищаем поле при фокусе, если там текст по умолчанию"""
        if self.context_text.get('1.0', tk.END).strip() == self.default_context.strip():
            self.context_text.delete('1.0', tk.END)
            self.context_text.configure(foreground='black')

    def on_context_focus_out(self, event):
        """Возвращаем текст по умолчанию, если поле пустое"""
        if not self.context_text.get('1.0', tk.END).strip():
            self.context_text.delete('1.0', tk.END)
            self.context_text.insert('1.0', self.default_context)
            self.context_text.configure(foreground='gray')

    def update_status(self, message):
        """Обновляет статус в статус-баре"""
        try:
            self.root.after(0, lambda: self.status_label.config(text=message))
        except Exception as e:
            self.log_event(f"Ошибка обновления статуса: {str(e)}")

    def log_event(self, message, level='info'):
        """Логирует событие и обновляет интерфейс"""
        if not hasattr(self, 'logger'):
            self.setup_logging()
        
        # Записываем в лог
        if level == 'error':
            self.logger.error(message)
        elif level == 'warning':
            self.logger.warning(message)
        else:
            self.logger.info(message)
        
        # Обновляем интерфейс через безопасный метод
        self.update_interface(message)

    def select_all(self, event=None):
        """Выделить весь текст"""
        self.context_text.tag_add(tk.SEL, "1.0", tk.END)
        self.context_text.mark_set(tk.INSERT, "1.0")
        self.context_text.see(tk.INSERT)
        return 'break'

    def copy_text(self, event=None):
        """Копировать текст"""
        if self.context_text.tag_ranges(tk.SEL):
            self.context_text.event_generate('<<Copy>>')
        return 'break'

    def cut_text(self, event=None):
        """Вырезать текст"""
        if self.context_text.tag_ranges(tk.SEL):
            self.context_text.event_generate('<<Cut>>')
        return 'break'

    def paste_text(self, event=None):
        """Вставить текст"""
        self.context_text.event_generate('<<Paste>>')
        return 'break'

    def show_context_menu(self, event):
        """Показать контекстное меню"""
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def create_presentation_guide(self, content, pdf_path, image_paths):
        """Создает PDF-гайд для презентации"""
        pdf_dir = os.path.dirname(pdf_path)
        pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
        guide_path = os.path.join(pdf_dir, f"{pdf_name}_presentation_guide.pdf")
        
        doc = SimpleDocTemplate(
            guide_path,
            pagesize=A4,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=72
        )
        
        # Создаем стили
        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(
            name='SlideNumber',
            fontSize=14,
            textColor=colors.HexColor('#1a237e'),
            spaceAfter=12
        ))
        styles.add(ParagraphStyle(
            name='MainPoint',
            fontSize=12,
            textColor=colors.HexColor('#000000'),
            spaceAfter=6,
            leading=14
        ))
        styles.add(ParagraphStyle(
            name='Details',
            fontSize=10,
            textColor=colors.HexColor('#424242'),
            spaceAfter=6,
            leading=12
        ))
        
        story = []
        
        # Добавляем заголовок
        title = f"Презентационный гайд: {pdf_name}"
        story.append(Paragraph(title, styles['Title']))
        story.append(Spacer(1, 20))
        
        # Для каждого слайда
        for i, (analysis, image_path) in enumerate(zip(content.split('• Слайд')[1:], image_paths), 1):
            if not analysis.strip():
                continue
            
            # Создаем мини-версию изображения слайда
            img = PILImage.open(image_path)
            img.thumbnail((200, 200))  # Уменьшаем размер
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='PNG')
            img_byte_arr = img_byte_arr.getvalue()
            
            # Добавляем номер слайда
            story.append(Paragraph(f"Слайд {i}", styles['SlideNumber']))
            
            # Добавляем миниатюру
            story.append(Image(io.BytesIO(img_byte_arr), width=2*inch, height=1.5*inch))
            story.append(Spacer(1, 10))
            
            # Разбираем и форматируем анализ
            parts = analysis.split('\n')
            for part in parts:
                if part.startswith('• ГЛАВНОЕ:'):
                    story.append(Paragraph(
                        f"<b>{part.replace('• ГЛАВНОЕ:', '🎯 Ключевой момент:')}</b>",
                        styles['MainPoint']
                    ))
                elif part.startswith('• ДЕТАЛИ:'):
                    story.append(Paragraph(
                        f"<b>{part.replace('• ДЕТАЛИ:', '💡 Акценты:')}</b>",
                        styles['Details']
                    ))
                elif part.startswith('• СВЯЗЬ:'):
                    story.append(Paragraph(
                        f"<i>{part.replace('• СВЯЗЬ:', 'Переход:')}</i>",
                        styles['Details']
                    ))
                elif part.strip() and not part.startswith('-'):
                    story.append(Paragraph(part, styles['Details']))
            
            story.append(Spacer(1, 20))
        
        # Создаем PDF
        doc.build(story)
        return guide_path

    def setup_logging(self):
        """Настраивает систему логирования"""
        # Создаем папку для логов если её нет
        log_dir = "logs"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        # Формируем имя файла лога с текущей датой
        log_file = os.path.join(log_dir, f"analysis_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
        
        # Настраиваем формат логирования
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s [%(levelname)s] %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        
        self.logger = logging.getLogger(__name__)
        self.logger.info("=== Запуск нового анализа ===")

    def update_interface(self, message):
        """Безопасное обновление интерфейса"""
        if hasattr(self, 'result_text'):
            self.root.after(0, lambda: self._safe_update_interface(message))

    def _safe_update_interface(self, message):
        """Выполняется в главном потоке"""
        try:
            self.result_text.insert(tk.END, message + "\n")
            self.result_text.see(tk.END)
        except Exception as e:
            print(f"Ошибка обновления интерфейса: {str(e)}")

    def on_closing(self):
        """Очистка при закрытии приложения"""
        try:
            # Очистка временных файлов
            if os.path.exists("slides_images"):
                for file in os.listdir("slides_images"):
                    try:
                        os.remove(os.path.join("slides_images", file))
                    except:
                        pass
                try:
                    os.rmdir("slides_images")
                except:
                    pass
            
            self.root.destroy()
        except Exception as e:
            self.log_event(f"Ошибка при закрытии приложения: {str(e)}", level='error')
            self.root.destroy()

    def test_analyze_first_10(self):
        """Тестовый анализ первых 10 слайдов"""
        if not self.file_path_var.get():
            self.show_error("Выберите PDF файл для анализа")
            return
        
        self.log_event("Начинаем тестовый анализ первых 10 слайдов")
        self.result_text.delete('1.0', tk.END)
        
        try:
            pdf_path = self.file_path_var.get()
            self.log_event(f"Анализируемый файл: {pdf_path}")
            
            # Создаем временную папку для слайдов
            output_folder = "test_slides"
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            
            # Конвертируем только первые 10 слайдов
            images = convert_from_path(pdf_path)[:10]
            image_paths = []
            
            for i, image in enumerate(images, 1):
                image_path = os.path.join(output_folder, f'slide_{i}.jpg')
                # Сохраняем с высоким качеством
                image.save(image_path, 'JPEG', quality=95)
                image_paths.append(image_path)
                self.log_event(f"Сохранен слайд {i}")
            
            # Анализируем каждый слайд
            for i, image_path in enumerate(image_paths, 1):
                self.log_event(f"\n=== Анализ слайда {i} ===")
                
                # Читаем изображение и конвертируем в base64
                try:
                    base64_image = self.encode_image_to_base64(image_path)
                    self.log_event(f"Слайд {i}: изображение успешно закодировано")
                except Exception as e:
                    self.log_event(f"Ошибка кодирования слайда {i}: {str(e)}", level='error')
                    continue
                
                # Формируем запрос к API
                try:
                    response = self.client.chat.completions.create(
                        model="gpt-4o-mini",
                        messages=[
                            {
                                "role": "user",
                                "content": [
                                    {
                                        "type": "text",
                                        "text": "Опишите, что вы видите на этом слайде презентации?"
                                    },
                                    {
                                        "type": "image_url",
                                        "image_url": {
                                            "url": f"data:image/jpeg;base64,{base64_image}"
                                        }
                                    }
                                ]
                            }
                        ],
                        max_tokens=300
                    )
                    
                    analysis = response.choices[0].message.content
                    self.log_event(f"\nРезультат анализа слайда {i}:")
                    self.log_event(analysis)
                    
                except Exception as e:
                    self.log_event(f"Ошибка при анализе слайда {i}: {str(e)}", level='error')
                
                # Пауза между запросами
                time.sleep(2)
            
            self.log_event("\nТестовый анализ завершен")
            
        except Exception as e:
            self.log_event(f"Общая ошибка тестового анализа: {str(e)}", level='error')
        
        finally:
            # Очищаем временные файлы
            try:
                for path in image_paths:
                    os.remove(path)
                os.rmdir(output_folder)
            except:
                pass

def main():
    try:
        check_dependencies()
        root = tk.Tk()
        app = BrandAnalyzerGUI(root)
        root.mainloop()
    except Exception as e:
        print(f"Ошибка при запуске приложения: {str(e)}")

if __name__ == "__main__":
    main() 