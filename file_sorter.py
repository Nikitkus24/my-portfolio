import pandas as pd
import numpy as np

# Создаем данные
data = {
    'Дата_продажи': pd.date_range(start='2026-01-01', periods=100, freq='D'),
    'Город': np.random.choice(['Москва', 'Санкт-Петербург', 'Казань', 'Новосибирск'], 100),
    'Товар': np.random.choice(['Ноутбук', 'Смартфон', 'Наушники', 'Монитор'], 100),
    'Количество': np.random.randint(1, 5, 100),
    'Цена': np.random.randint(1000, 50000, 100),
    'Скидка_%': np.random.choice([0, 5, 10, 15], 100)
}
df = pd.DataFrame(data)

# Добавим немного "грязных" данных для реализма
df.loc[5:10, 'Город'] = np.nan  # Пропуски
df.loc[20:22, 'Цена'] = np.nan

# Сохраняем в Excel
df.to_excel('test_sales_russian.xlsx', index=False)
print("Файл 'test_sales_russian.xlsx' создан!")