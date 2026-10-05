import os
from pathlib import Path

# Папка для тестов
TEST_DIR = Path(r"C:\Test")

def create_tests():
    # Создаем структуру папок
    folder1 = TEST_DIR / "Folder1"
    folder2 = TEST_DIR / "Folder2"
    subfolder = folder1 / "SubFolder"
    
    for folder in [folder1, folder2, subfolder]:
        folder.mkdir(parents=True, exist_ok=True)
        
    # 1. Текстовый файл оригинал
    doc1_content = "Это тестовый документ для проверки поиска дубликатов.\nСодержимое 100% совпадает."
    (folder1 / "document1.txt").write_text(doc1_content, encoding="utf-8")
    
    # 2. Точная копия оригинала (Дубликат по хэшу) в другой папке
    (folder2 / "copy_of_document1.txt").write_text(doc1_content, encoding="utf-8")
    
    # 3. Уникальный текстовый файл
    (folder2 / "unique_note.txt").write_text("Совершенно другой текст, которого больше нет.", encoding="utf-8")
    
    # 4. Файл в подпапке
    (subfolder / "deep_file.log").write_text("2026-10-05 LOG: Everything operates normally.", encoding="utf-8")
    
    # 5. Простой CSV файл
    csv_content = "id;name;price\n1;Laptop;1000\n2;Mouse;25"
    (folder1 / "data.csv").write_text(csv_content, encoding="utf-8")

    print(f"✅ Тестовая структура успешно создана в '{TEST_DIR}'!")
    print("Созданы файлы:")
    print(" - Folder1\\document1.txt")
    print(" - Folder2\\copy_of_document1.txt (Дубликат document1)")
    print(" - Folder2\\unique_note.txt")
    print(" - Folder1\\SubFolder\\deep_file.log")
    print(" - Folder1\\data.csv")

if __name__ == "__main__":
    create_tests()