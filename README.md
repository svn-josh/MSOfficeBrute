# MSOfficeBrute

🚀 A fast, parallel brute-force cracker for password-protected Microsoft Office files (`.xlsx`, `.docx`, etc.).  
Leverage your CPU's full power, customize the character set, and crack passwords through raw processing power.

## 🔧 Features

- Supports `.docx`, `.xlsx`, `.pptx` (modern Office formats)
- Multi-core support via Python multiprocessing
- Real-time progress display using `tqdm`
- Configurable character set (lowercase, uppercase, digits, symbols)
- Adjustable password length and chunk size
- Automatically detects number of logical CPU cores

## ⚙️ Requirements

- Python 3.10+
- [msoffcrypto-tool](https://pypi.org/project/msoffcrypto-tool/)
- tqdm

Install dependencies:

```bash
pip install msoffcrypto-tool tqdm
```

## 🚀 Usage

Edit the configuration block at the top of `main.py`:

```python
file_path = "protected_file.xlsx"
min_length = 1
max_length = 4
chunk_size = 5000
```

Enable or disable character types:

```python
use_lowercase = True
use_uppercase = False
use_digits = True
use_symbols = False
```

Then simply run the script:

```bash
python main.py
```

You'll see output like:

```
🚀 Starting brute-force attack
⚙️  Processes: 24
🔤 Charset: abcdefghijklmnopqrstuvwxyz0123456789
🔢 Password length: 1 to 4
🧩 Chunk size: 5000

✅ Password found: mypass123
⏱️  Finished in 58.43 seconds.
```

## ❗ Disclaimer

This tool is intended for educational purposes or authorized testing only.  
**Do not use it on files or systems you do not own or have permission to test.**

## 📄 License

MIT License
