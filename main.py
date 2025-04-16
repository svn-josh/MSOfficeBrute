import multiprocessing
import msoffcrypto
import io
import time
import itertools
import string
from tqdm import tqdm

# === Konfiguration ===
file_path = "geschuetzte_datei.xlsx"  # Pfad zur geschützten Datei
min_length = 4
max_length = 4
chunk_size = 5000
process_count = multiprocessing.cpu_count()

# === Zeichensatz Optionen ===
use_lowercase = True
use_uppercase = False
use_digits = False
use_symbols = False

# === Zeichensatz erstellen ===
characters = ""
if use_lowercase:
    characters += string.ascii_lowercase
if use_uppercase:
    characters += string.ascii_uppercase
if use_digits:
    characters += string.digits
if use_symbols:
    characters += string.punctuation

if not characters:
    raise ValueError("⚠️ Kein Zeichensatz ausgewählt! Bitte aktiviere mindestens eine Option.")

# === Passwort-Generator ===
def generate_passwords():
    for length in range(min_length, max_length + 1):
        for pwd in itertools.product(characters, repeat=length):
            yield ''.join(pwd)

# === In Stücke aufteilen ===
def chunked_iterable(iterable, size):
    it = iter(iterable)
    while True:
        chunk = list(itertools.islice(it, size))
        if not chunk:
            return
        yield chunk

# === Passwort testen ===
def check_password(password):
    try:
        decrypted = io.BytesIO()
        with open(file_path, "rb") as file:
            office_file = msoffcrypto.OfficeFile(file)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
        return True, password
    except Exception:
        return False, password

# === Worker Prozess ===
def worker(passwords_chunk, found_flag, counter, locks, progress_queue):
    counter_lock, found_flag_lock = locks
    for password in passwords_chunk:
        # Frühzeitiger Abbruch wenn Passwort gefunden
        with found_flag_lock:
            if found_flag.value:
                return None
        success, tested_password = check_password(password)
        with counter_lock:
            counter.value += 1
        progress_queue.put(1)  # Send progress update to the main process
        if success:
            with found_flag_lock:
                found_flag.value = 1
            return tested_password
    return None

# === Fortschrittsanzeige aktualisieren ===
def update_progress_bar(progress_queue):
    while True:
        progress = progress_queue.get()
        if progress is None:
            break
        yield progress

# === Fortschrittsanzeige Thread ===
def progress_updater(progress_queue, pbar):
    while True:
        progress = progress_queue.get()
        if progress is None:
            break
        pbar.update(progress)

if __name__ == "__main__":
    print("\n🚀 Starte Brute-Force Angriff")
    print(f"⚙️  Prozesse: {process_count}")
    print(f"🔤 Zeichensatz: {characters}")
    print(f"🔢 Passwortlänge: {min_length} bis {max_length}")
    print(f"🧩 Chunk-Größe: {chunk_size}\n")

    start_time = time.time()

    manager = multiprocessing.Manager()
    counter = manager.Value('i', 0)
    found_flag = manager.Value('i', 0)
    counter_lock = manager.Lock()
    found_flag_lock = manager.Lock()

    locks = (counter_lock, found_flag_lock)
    progress_queue = manager.Queue()

    total_passwords = sum(len(characters) ** length for length in range(min_length, max_length + 1))
    password_chunks = list(chunked_iterable(generate_passwords(), chunk_size))

    with multiprocessing.Pool(process_count) as pool, tqdm(total=total_passwords) as pbar:
        results = []

        # Start a thread to update the progress bar in the main process
        from threading import Thread
        progress_thread = Thread(target=progress_updater, args=(progress_queue, pbar))
        progress_thread.start()

        for chunk in password_chunks:
            if found_flag.value:
                break
            result = pool.apply_async(worker, args=(chunk, found_flag, counter, locks, progress_queue))
            results.append((result, len(chunk)))

        # Ergebnisse auslesen
        found_password = None
        for result, chunk_length in results:
            password = result.get()
            if password:
                found_password = password
                break

        # Signal the progress updater to stop
        progress_queue.put(None)
        progress_thread.join()

        pool.terminate()
        pool.join()

    if found_password:
        print(f"\n✅ Passwort gefunden: {found_password}")
    else:
        print("\n❌ Passwort wurde nicht gefunden.")

    elapsed_time = time.time() - start_time
    print(f"⏱️  Fertig in {elapsed_time:.2f} Sekunden.")
