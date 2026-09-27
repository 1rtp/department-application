# Система управління кафедрою

Десктопний WinForms застосунок для комплексного управління персоналом, академічними дисциплінами, науковими дослідженнями та звітністю кафедри комп'ютерних наук та інформаційних технологій.

---

## Технологічний стек

| Частина | Технології |
|---|---|
| **Язык** | C# (.NET Framework 4.7.2) |
| **UI Framework** | Windows Forms |
| **База даних** | SQLite |
| **Звітність** | FastReport .NET |
| **Сеціалізація даних** | Власна бібліотека (SerializerLib), JSON, XML, Binary |
| **Office Integration** | MS Office Interop (Excel, Word), ClosedXML, DocX, Spire.Doc |
| **Логування** | NLog 6.0.5 |

---

## Структура проєкту

```
КП Кафедра/                      # Основний проєкт WinForms
├── Forms/                        # UI форми
│   ├── FormLogin.cs             # Форма входу
│   ├── FormAssignment.cs        # Управління навантаженням викладачів
│   ├── FormTeacher.cs           # Управління викладачами
│   ├── FormSubject.cs           # Управління дисциплінами
│   ├── FormResearch.cs          # Управління науковими дослідженнями
│   ├── FormParticipation.cs     # Управління участю в проєктах
│   ├── FormReport.cs            # Генерація звітів
│   ├── FormReportViewer.cs      # Перегляд звітів
│   ├── FormTables.cs            # Перегляд таблиць даних
│   ├── FormSettings.cs          # Налаштування системи
│   └── ToastForm.cs             # Сповіщення користувача
├── Program.cs                    # Entry point, моделі даних (Teacher, Subject, Research, etc.)
├── DatabaseSQL.cs               # Ініціалізація та робота з БД
├── DataService.cs               # Сервіс отримання/збереження даних
├── AppSettings.cs               # Конфігурація застосунку
├── FormMainMenu.cs              # Головне меню
├── LanguageManager.cs           # Управління мовами (UK/EN)
├── IteratorPattern.cs           # Паттерн Iterator для колекцій
├── RepCOM.cs                    # Інтеграція з COM об'єктами
├── ReportBridge.cs              # Мост до FastReport
├── RoundButton.cs               # Кастомний компонент кнопки
├── DatePicker.cs                # Кастомний календар для вибору дати
├── NLog.config                  # Конфіг логування
├── App.config                   # Конфіг застосунку
├── Resources/                    # Ресурси (іконки, локалізація)
│   ├── Strings.resx            # Текстові ресурси (укр.)
│   ├── Strings.en.resx         # Текстові ресурси (англ.)
│   └── *.png                    # Іконки (icons8)
├── Data/                         # Файли даних
│   ├── department.db            # SQLite база даних
│   ├── init_data.sql            # Скрипт ініціалізації БД
│   ├── department.json          # Експорт даних (JSON)
│   ├── department.xml           # Експорт даних (XML) для FastReport
│   └── department.bin           # Експорт даних (Binary)
├── Reports/                      # Шаблони звітів FastReport
└── КП Кафедра.csproj           # Конфіг проєкту

SerializerLib/                    # Окремий проєкт — бібліотека сеціалізації
├── SerializerLib.csproj        # Конфіг бібліотеки
└── ...                          # Класи для серіалізації

КП Кафедра.sln                   # Visual Studio рішення
```

---

## Основні функції

### 👨‍🏫 Управління викладачами
- Реєстрація та редагування даних викладача (ПІБ, посада, контакти)
- Отримання та звільнення з кафедри
- Прив'язка до спеціальностей

### 📚 Управління дисциплінами
- Додавання та редагування дисциплін
- Розподіл за семестрами та спеціальностями
- Облік годин та статусу

### 👥 Навантаження викладачів
- Розподіл завдань (Assignment) між викладачами
- Облік планових та фактичних годин
- Типи занять (лекції, практики, семінари)

### 🔬 Наукові дослідження
- Реєстрація проєктів та дослідницьких напрямків
- Управління участю викладачів в науковій роботі
- Облік часових рамок (дата початку / завершення)

### 📊 Звітність
- Генерація звітів FastReport у форматі PDF, Excel, Word
- Експорт даних у JSON, XML, Binary
- Користувацька типографія та структурування даних

### 🌐 Багатомовність
- Підтримка української та англійської мов
- Динамічна зміна мови без перезавантаження
- Локалізовані ресурси (Strings.resx, Strings.en.resx)

---

## Швидкий старт

### Передумови

- **Visual Studio 2019+** (або VS Code з .NET Framework інструментами)
- **.NET Framework 4.7.2** (повинен бути встановлено)
- **Windows OS** (невід'ємна вимога для Windows Forms)

### 1. Клонування репозиторію

```bash
git clone https://github.com/1rtp/department-application.git
cd department-application
```

### 2. Відновлення пакетів NuGet

У Visual Studio:

```
Tools → NuGet Package Manager → Package Manager Console
```

Виконайте:

```powershell
# Для головного проєкту
Update-Package

# Для SerializerLib
cd ..\SerializerLib
Update-Package
cd ..\
```

Або натисніть на рішення правою кнопкою → **Restore NuGet Packages**.

### 3. Збирання проєкту SerializerLib

1. У **Solution Explorer** клацніть правою кнопкою на проєкт **SerializerLib**
2. Виберіть **Build**
3. Дочекайтесь завершення без помилок

> ⚠️ **Важливо**: Головний проєкт залежить від скомпільованої DLL з SerializerLib.
> Якщо ви отримаєте помилку про відсутність бібліотеки, перебудуйте SerializerLib.

### 4. Збирання та запуск головного проєкту

1. Клацніть на рішення **КП Кафедра** (проєкт з іконкою WinForms)
2. Натисніть **Build → Build КП Кафедра**
3. Запустіть проєкт **F5** або **Debug → Start Debugging**

---

## Використання

### Дані для входу

При першому запуску використовуйте **стандартні облікові дані адміністратора**:

| Поле | Значення |
|---|---|
| **Email** | `romaasericyn@gmail.com` |
| **Password** | `123456789` |

Ці дані зберігаються у файлі:

```
КП Кафедра/
└── Data/
    └── admin.json
```

Ви можете змінити їх прямо у файлі або через форму налаштувань.

### Основні операції

#### Додавання викладача

1. Головне меню → **Викладачі**
2. Натисніть **Додати нового**
3. Заповніть дані (ПІБ, посада, контакти, спеціальність)
4. Натисніть **Зберегти**

#### Додавання дисципліни

1. Головне меню → **Дисципліни**
2. Натисніть **Додати**
3. Виберіть спеціальність, введіть назву, семестр, кількість годин
4. Натисніть **OK**

#### Розподіл навантаження

1. Головне меню → **Навантаження**
2. Виберіть викладача та дисципліну
3. Задайте планові години, тип заняття
4. Натисніть **Зберегти**

#### Генерація звіту

1. Головне меню → **Звіти**
2. Виберіть тип звіту (за викладачами, дисциплінами, дослідженнями)
3. Налаштуйте параметри
4. Натисніть **Сформувати**
5. Оберіть формат експорту (PDF, Excel, Word)

### Робота з даними FastReport

Перед формуванням звіту у FastReport:

1. **Оновіть XML-файл**: Головне меню → **Налаштування** → **Експортувати дані** → **XML**
2. Це забезпечить, що звіти містять актуальні дані з бази
3. Шаблони звітів розташовані у папці **Reports/**

### Експорт даних

Застосунок дозволяє експортувати весь датасет у три формати:

| Формат | Розташування | Призначення |
|---|---|---|
| **JSON** | `Data/department.json` | Обмін даними, веб-інтеграція |
| **XML** | `Data/department.xml` | Джерело даних для FastReport |
| **Binary** | `Data/department.bin` | Компактне зберігання, швидкість |

---

## Структура бази даних

SQLite база `department.db` містить такі таблиці:

| Таблиця | Опис |
|---|---|
| `Specialties` | Спеціальності кафедри |
| `Teachers` | Викладачі з контактами та статусом |
| `Subjects` | Академічні дисципліни |
| `LessonTypes` | Типи занять (лекція, практика, семінар) |
| `Assignments` | Навантаження викладачів |
| `Researches` | Наукові проєкти |
| `Participations` | Участь викладачів у дослідженнях |
| `Users` | Користувачі системи (автентифікація) |

Схема БД ініціалізується автоматично з файлу **`init_data.sql`** при першому запуску.

---

## Налаштування

### Мова інтерфейсу

1. Головне меню → **Налаштування**
2. Виберіть мову: **Українська** або **English**
3. Застосунок перезавантажиться автоматично

Налаштування зберігаються в `App.config`:

```xml
<appSettings>
    <add key="Language" value="Ukrainian" />
</appSettings>
```

### Конфіг логування (NLog)

Файл `КП Кафедра/NLog.config` керує записуванням логів:

```xml
<target name="file" xsi:type="File" 
  fileName="Logs/${shortdate}.log" />
```

Логи писаються у папку **Logs/** в корені проєкту.

---

## Особливості реалізації

- **Паттерни проектування**: Iterator (для перебору колекцій), Bridge (для FastReport), Singleton (для DataService)
- **COM-інтеграція**: Безпосередня робота з Excel та Word через MS Office Interop
- **Асинхронні операції**: Експорт великих датасетів в окремих потоках
- **Валідація**: Перевірка даних на рівні форм та БД
- **Логування**: Всі операції логуються через NLog для діагностики проблем

---

## Поширені проблеми

| Проблема | Причина | Рішення |
|---|---|---|
| *"SerializerLib не знайдено"* | DLL не збудована | Перебудуйте проєкт SerializerLib, переконайтеся шляху в .csproj |
| *Помилка при запуску FastReport* | Відсутні файли звітів | Перевірте папку `Reports/` на наявність `.frx` файлів |
| *База даних не ініціалізується* | Відсутній `init_data.sql` | Перевірте папку `Data/` та шляхи в `Program.cs` |
| *Не вдається підключитися до бази* | Неправильний шлях або доступ | Переконайтеся, що папка `Data/` існує і доступна для запису |
| *Office Interop викидає помилку* | MS Office не встановлений | Встановіть Microsoft Office або скористайтесь альтернативою (ClosedXML, DocX) |
| *Мова не змінюється* | Помилка в локалізації ресурсів | Перевірте файли `Strings.resx` та `Strings.en.resx` |

---

## Розробка

### Додавання нової форми

1. У папці `Forms/` створіть новий WinForms клас: **Add → New Item → Windows Form**
2. Назвіть його `FormYourName.cs`
3. На главному меню додайте кнопку з обробником:

```csharp
private void btnYourForm_Click(object sender, EventArgs e)
{
    FormYourName form = new FormYourName();
    form.ShowDialog();
}
```

### Додавання нової таблиці до БД

1. Відредагуйте файл `Data/init_data.sql` і додайте `CREATE TABLE` запит
2. Оновіть клас `DatabaseSQL.cs` з новими методами для CRUD операцій
3. Перестартуйте застосунок — база буде заново ініціалізована

### Локалізація

Додайте новий текст:

1. Відкрийте `Resources/Strings.resx` (укр.) та `Strings.en.resx` (англ.)
2. Додайте новий ресурс з унікальним ключем (напр., `lblNewText`)
3. У коді використовуйте:

```csharp
label.Text = Resources.Strings.lblNewText;
```

---

## Автор

**Ящеріцин Роман Євгенович**

---

## Ліцензія

Вказати ліцензію (якщо є)

---

# Department Management System

Desktop WinForms application for comprehensive management of personnel, academic disciplines, scientific research, and reporting for the Computer Science and Information Technologies Department.

---

## Technology Stack

| Component | Technologies |
|---|---|
| **Language** | C# (.NET Framework 4.7.2) |
| **UI Framework** | Windows Forms |
| **Database** | SQLite |
| **Reporting** | FastReport .NET |
| **Data Serialization** | Custom library (SerializerLib), JSON, XML, Binary |
| **Office Integration** | MS Office Interop (Excel, Word), ClosedXML, DocX, Spire.Doc |
| **Logging** | NLog 6.0.5 |

---

## Quick Start

### Prerequisites

- **Visual Studio 2019+** or **Visual Studio Code with .NET Framework tools**
- **.NET Framework 4.7.2** installed
- **Windows OS** (required for Windows Forms)

### Setup Steps

1. **Clone repository**:
   ```bash
   git clone https://github.com/1rtp/department-application.git
   cd department-application
   ```

2. **Restore NuGet packages**:
   - Open solution in Visual Studio
   - Right-click solution → **Restore NuGet Packages**

3. **Build SerializerLib**:
   - Right-click **SerializerLib** project → **Build**
   - Ensure build completes without errors

4. **Build and run main project**:
   - Select **КП Кафедра** project
   - Press **F5** or **Debug → Start Debugging**

### Default Login Credentials

| Field | Value |
|---|---|
| **Email** | `romaasericyn@gmail.com` |
| **Password** | `123456789` |

---

## Main Features

- **Teacher Management** — registration, editing, assignment to specialties
- **Academic Discipline Management** — subjects by semester and specialty
- **Workload Distribution** — assign lessons and teaching hours to teachers
- **Scientific Research** — manage projects and teacher participation
- **Report Generation** — create and export reports in PDF, Excel, Word formats
- **Data Export** — JSON, XML, Binary formats
- **Multilingual Support** — Ukrainian and English
- **NLog Integration** — comprehensive logging and diagnostics

---

## Project Structure

The main application **КП Кафедра** contains:
- **Forms/** — WinForms UI classes
- **Data/** — SQLite database and initial data script
- **Reports/** — FastReport templates
- **Resources/** — localization strings and icons

**SerializerLib** — separate DLL project for data serialization

---

## Important Notes

⚠️ Ensure all projects in the solution build successfully before running.
⚠️ The `Data` folder must remain in its original location relative to the executable.
⚠️ FastReport templates require updated XML export file for current data.

For detailed documentation in Ukrainian, see the top section of this README.
