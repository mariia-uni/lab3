from __future__ import annotations

import warnings
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style, apply openpyxl's default",
    category=UserWarning,
    module="openpyxl.styles.stylesheet",
)


@dataclass
class TxnInfo:
    payment_due_date: str = ""  # A2
    industry_category: str = ""  # B2
    credit_card_number: str = ""  # C2
    merchant_category: str = ""  # D2
    card_amount: float = 0.0  # E2
    card_currency: str = ""  # F2
    txn_amount: float = 0.0  # G2
    txn_currency: str = ""  # H2
    ending_balance: float = 0.0  # I2
    balance_currency: str = ""  # J2

    @classmethod
    def from_row(cls, row: list) -> "TxnInfo":
        def safe_float(value):
            if value is None or pd.isna(value):
                return 0.0
            return float(value)

        return cls(
            payment_due_date=str(row[0]) if len(row) > 0 and row[0] is not None and not pd.isna(row[0]) else "",
            industry_category=str(row[1]) if len(row) > 1 and row[1] is not None and not pd.isna(row[1]) else "",
            credit_card_number=str(row[2]) if len(row) > 2 and row[2] is not None and not pd.isna(row[2]) else "",
            merchant_category=str(row[3]) if len(row) > 3 and row[3] is not None and not pd.isna(row[3]) else "",
            card_amount=safe_float(row[4]) if len(row) > 4 else 0.0,
            card_currency=str(row[5]) if len(row) > 5 and row[5] is not None and not pd.isna(row[5]) else "",
            txn_amount=safe_float(row[6]) if len(row) > 6 else 0.0,
            txn_currency=str(row[7]) if len(row) > 7 and row[7] is not None and not pd.isna(row[7]) else "",
            ending_balance=safe_float(row[8]) if len(row) > 8 else 0.0,
            balance_currency=str(row[9]) if len(row) > 9 and row[9] is not None and not pd.isna(row[9]) else "",
        )


@dataclass
class CreditCardStatement:
    statement_period: str = ""  # A1:J1, e.g. "Історія операцій за період 28.09.2026 - 28.10.2026"
    txn: list[TxnInfo] = field(default_factory=list)

    @classmethod
    def from_excel(cls, file_path: str | Path) -> "CreditCardStatement":
        path = Path(file_path)
        sheet = pd.read_excel(path, sheet_name="Виписки", header=None)

        if sheet.empty:
            raise ValueError(f"Файл {path} порожній.")

        statement_period = ""
        if not sheet.empty and not pd.isna(sheet.iloc[0, 0]):
            statement_period = str(sheet.iloc[0, 0]).strip()

        header_row_index = None
        for idx, row in sheet.iterrows():
            first_cell = row.iloc[0] if len(row) > 0 else ""
            if isinstance(first_cell, str) and "Дата" in first_cell:
                header_row_index = idx
                break

        if header_row_index is None:
            rows_to_parse = sheet.iloc[2:]
        else:
            rows_to_parse = sheet.iloc[header_row_index + 1 :]

        transactions = []
        for row in rows_to_parse.itertuples(index=False, name=None):
            if len(row) < 10:
                continue
            if all(item is None or (isinstance(item, float) and pd.isna(item)) for item in row[:10]):
                continue
            transactions.append(TxnInfo.from_row(list(row)))

        return cls(statement_period=statement_period, txn=transactions)

    def total_spent(self) -> float:
        return sum(abs(t.card_amount) for t in self.txn if t.card_amount < 0)

    def total_income(self) -> float:
        return sum(t.card_amount for t in self.txn if t.card_amount > 0)

    def category_spending_summary(self) -> dict[str, float]:
        summary: dict[str, float] = {}
        for item in self.txn:
            if item.card_amount < 0:
                category = item.industry_category.strip() or "Інше"
                summary[category] = summary.get(category, 0.0) + abs(item.card_amount)
        return dict(sorted(summary.items(), key=lambda x: x[1], reverse=True))

    def top_expense_category(self) -> tuple[str, float]:
        summary = self.category_spending_summary()
        if not summary:
            return ("Немає даних", 0.0)
        category, amount = max(summary.items(), key=lambda item: item[1])
        return category, amount

    def top_expense_merchants(self, limit: int = 5) -> list[tuple[str, float]]:
        expenses: dict[str, float] = {}
        for item in self.txn:
            if item.card_amount < 0:
                merchant = item.merchant_category.strip() or "Невідомо"
                expenses[merchant] = expenses.get(merchant, 0.0) + abs(item.card_amount)

        return sorted(expenses.items(), key=lambda item: item[1], reverse=True)[:limit]


def print_menu() -> None:
    print("\n=== Меню аналізу витрат ===")
    print("1. Загальна сума витрат")
    print("2. Категорія з найбільшими витратами")
    print("3. Найбільші витрати за магазинами/операціями")
    print("4. Показати всі категорії витрат")
    print("5. Показати загальний дохід")
    print("0. Вийти")


def handle_user_menu(statement: CreditCardStatement) -> None:
    while True:
        print_menu()
        choice = input("Оберіть дію: ").strip()

        if choice == "1":
            print(f"\nЗагальна сума витрат за період: {statement.total_spent():.2f} UAH")
        elif choice == "2":
            category, total = statement.top_expense_category()
            print(f"\nКатегорія з найбільшими витратами: {category}")
            print(f"Сума: {total:.2f} UAH")
        elif choice == "3":
            merchants = statement.top_expense_merchants()
            print("\nТоп витрат за магазинами/операціями:")
            if not merchants:
                print("Немає даних про витрати.")
            else:
                for index, (merchant, amount) in enumerate(merchants, start=1):
                    print(f"{index}. {merchant}: {amount:.2f} UAH")
        elif choice == "4":
            summary = statement.category_spending_summary()
            print("\nВитрати по категоріях:")
            if not summary:
                print("Немає даних про витрати.")
            else:
                for category, amount in summary.items():
                    print(f"- {category}: {amount:.2f} UAH")
        elif choice == "5":
            print(f"\nЗагальний дохід за період: {statement.total_income():.2f} UAH")
        elif choice == "0":
            print("\nДо побачення!")
            break
        else:
            print("\nНевірний вибір. Спробуйте ще раз.")


def main() -> None:
    try:
        file_path = input("Введіть шлях до файлу з історією транзакцій (.xlsx): ").strip()
        if not file_path:
            file_path = str(Path(__file__).with_name("transactions.xlsx"))

        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"Файл не знайдено: {path}")
        if path.suffix.lower() != ".xlsx":
            raise ValueError(f"Некоректний формат файлу: {path.suffix or 'без розширення'}")

        statement = CreditCardStatement.from_excel(path)

        print(f"\nПеріод: {statement.statement_period}")
        print(f"Сума витрат за період: {statement.total_spent():.2f} UAH")
        handle_user_menu(statement)
    except FileNotFoundError as exc:
        print("\n[Помилка] Файл не знайдено.")
        print(f"Шлях: {exc}")
        print("Перевірте, чи вказано правильний шлях до .xlsx-файлу, і спробуйте ще раз.\n")
    except ValueError as exc:
        print("\n[Помилка] Невірний формат файлу.")
        print(f"Деталі: {exc}")
        print("Потрібен файл у форматі .xlsx з історією транзакцій.\n")
    except KeyboardInterrupt:
        print("\n\nВиконання програми перервано користувачем.")


if __name__ == "__main__":
    main()