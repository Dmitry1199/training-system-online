TRAINING SYSTEM ONLINE v1

Окрема копія робочої v4.7. Локальну робочу версію не змінює.

Локально:
1) Відкрити PowerShell у цій папці.
2) python -m pip install -r requirements.txt
3) python run_online.py
4) Відкрити http://127.0.0.1:5000

Для Render/іншого production-сервера:
gunicorn --bind 0.0.0.0:$PORT training_web_v4_7:app

ВАЖЛИВО: ця перша ONLINE-збірка зберігає існуючу Excel-логіку v4.7.
Міграцію результатів у PostgreSQL робимо окремим кроком після перевірки,
щоб не ризикувати історією.
