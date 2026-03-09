<div align="center">

<img src="https://img.icons8.com/3d-fluency/128/iphone.png" width="100" alt="iPhone">

# 📱 iPhone Media Browser

### Просмотр фото и видео с iPhone на Windows — без iTunes, без копирования всего подряд

[![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6?style=for-the-badge&logo=windows&logoColor=white)](https://microsoft.com)
[![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)](LICENSE)
[![Version](https://img.shields.io/badge/Version-0.0.5-orange?style=for-the-badge)]()

---

**Подключил iPhone → Нажал кнопку → Смотришь фото и видео**

</div>

<br>

## 🎯 Зачем это нужно?

> Windows видит iPhone как **MTP-устройство** — нельзя просто открыть папку и работать с файлами.
> Проводник медленный, не показывает превью, а iTunes — это iTunes 😩

| Проблема | Решение |
|:---------|:--------|
| 📂 Файлы разбросаны по 100+ папкам | Автоматический поиск **всех** медиафайлов |
| 🐌 Проводник тормозит на MTP | Быстрое сканирование с прогрессом и ETA |
| 🖼️ Нет превью в проводнике | Предпросмотр фото и видео в приложении |
| 📋 Нет сортировки | Сортировка по размеру, дате, типу, имени |
| 🔍 Не найти нужный файл | Поиск и фильтры (фото / видео / AAE) |
| 📦 HEIC не открывается | Встроенная поддержка HEIC/HEIF |

<details>
<summary><b>🏷️ v0.0.5 — «Финальный UI»</b> (текущая версия)</summary>

**Новое:**
- ✅ Бегущая полоса (indeterminate) на шагах 1–2 — видно что программа работает
- ✅ Точный % (determinate) на шаге 3 — видно сколько осталось
- ✅ Живой таймер ⏱️ — обновляется каждые 0.5 сек
- ✅ Крупный номер шага — «ШАГ 2 из 3 — Чтение структуры»
- ✅ ETA — «⏱️ Осталось: ~1.4 мин»
- ✅ Колбэк на каждый файл — счётчик обновляется в реальном времени
- ✅ Автоскрытие прогресс-панели после завершения
</details>

<details>
<summary><b>🏷️ v0.0.4 — «Живой прогресс»</b></summary>

**Новое:**
- ✅ Пошаговое обнаружение: Корень → Storage → DCIM → 100APPLE…
- ✅ Колбэк на КАЖДЫЙ элемент — UI обновляется в реальном времени
- ✅ Три чётких шага: Поиск → Структура → Файлы
- ❌ Полоса на 0% — кажется что не работает
- ❌ Нет таймера — непонятно сколько прошло
</details>

## ⚡ Быстрый старт

**1️⃣ Установка зависимостей**

```bash
pip install pywin32 Pillow pillow-heif opencv-python

**2️⃣ Подключите iPhone**

🔌 USB-кабель → 🔓 Разблокируйте экран → ✅ «Доверять этому компьютеру» → 📦 Нужен iTunes или Apple Devices

3️⃣ Запуск

Bash

python iphone_browser.py
<br>

