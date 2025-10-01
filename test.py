import pandas as pd
from pathlib import Path
import json
from typing import Dict, List, Any
from datetime import datetime
import os
from allpairspy import AllPairs
import string
import random
import uuid

def clear_console():
    """Очищает консоль в зависимости от операционной системы"""
    os.system('cls' if os.name == 'nt' else 'clear')

def create_attributes_description_excel(file_path: Path, sheet_name: str, headers: List[str], examples: List[str], 
                                      field_types: Dict[str, str], min_values: Dict[str, str], 
                                      max_values: Dict[str, str], maxlen_values: Dict[str, str]) -> Path:
    """Создает отдельный Excel файл с описанием атрибутов для конкретного листа"""
    output_dir = file_path.parent / file_path.stem
    output_dir.mkdir(exist_ok=True)
    
    description_data = []
    test_parameters = []
    
    for header, example in zip(headers, examples):
        field_type = field_types.get(header, '').lower()
        min_val = min_values.get(header, '')
        max_val = max_values.get(header, '')
        maxlen_val = maxlen_values.get(header, '')
        
        # Получаем только имя поля (последнюю часть пути)
        field_name = header.split('.')[-1].replace('[0]', '')
        
        # Формируем значения для тестирования
        test_values = []
        
        # Разделяем значения из столбца "Пример" по символу ";"
        example_values = [ex.strip() for ex in example.split(';')] if example and ';' in example else [example] if example else []
        
        # Для строковых типов (не массивов)
        if any(t in field_type for t in ['string', 'text']) and 'array' not in field_type:
            # Добавляем все значения из примера как отдельные
            for ex in example_values:
                if ex and ex not in test_values:
                    test_values.append(ex)
            
            if maxlen_val:
                try:
                    max_len = int(maxlen_val)
                    # Добавляем строку минимальной длины (1 символ)
                    if 'a' not in test_values:
                        test_values.append('a')
                    # Добавляем строку максимальной длины
                    if max_len > 1:
                        long_string = 'a' * max_len
                        if long_string not in test_values:
                            test_values.append(long_string)
                except ValueError:
                    pass
        
        # Для числовых типов (не массивов)
        elif any(t in field_type for t in ['number', 'integer', 'int', 'float', 'double', 'decimal']) and 'array' not in field_type:
            # Пробуем преобразовать примеры в числа
            for ex in example_values:
                try:
                    num = float(ex) if '.' in ex else int(ex)
                    if str(num) not in test_values:
                        test_values.append(str(num))
                except (ValueError, TypeError):
                    pass
            
            if min_val and min_val not in test_values:
                try:
                    min_num = float(min_val) if '.' in min_val else int(min_val)
                    test_values.append(str(min_num))
                except ValueError:
                    pass
            if max_val and max_val not in test_values:
                try:
                    max_num = float(max_val) if '.' in max_val else int(max_val)
                    test_values.append(str(max_num))
                except ValueError:
                    pass
        
        # Для булевых типов
        elif any(t in field_type for t in ['boolean', 'bool']):
            for ex in example_values:
                if ex.lower() in ['true', 'false', '1', '0', 'y', 'n']:
                    bool_val = 'true' if ex.lower() in ['true', '1', 'y'] else 'false'
                    if bool_val not in test_values:
                        test_values.append(bool_val)
            if 'true' not in test_values:
                test_values.append('true')
            if 'false' not in test_values:
                test_values.append('false')
        
        # Для массивов
        elif 'array' in field_type:
            if example:
                try:
                    if example.startswith('[') and example.endswith(']'):
                        parsed = json.loads(example)
                        if isinstance(parsed, list):
                            test_values.append(parsed)
                    else:
                        # Разделяем примеры по ";" для массивов строк
                        if 'string' in field_type:
                            test_values.append(example_values)
                        else:
                            test_values.append(example_values)
                except:
                    test_values.append(example_values)
            
            if 'string' in field_type and maxlen_val:
                try:
                    max_len = int(maxlen_val)
                    test_values.append(['a'])
                    if max_len > 1:
                        long_string = 'a' * max_len
                        test_values.append([long_string])
                except ValueError:
                    pass
        
        # Если нет значений, добавляем значение по умолчанию
        if not test_values:
            test_values.append(get_default_value(field_type))
        
        # Сохраняем параметры для расчета попарных тестов
        test_parameters.append(test_values)
        
        # Форматируем значения для вывода
        if test_values:
            if (any(t in field_type for t in ['string', 'text']) and 'array' not in field_type) or \
               ('array' in field_type and 'string' in field_type):
                formatted_values = []
                for val in test_values:
                    if isinstance(val, list):
                        lengths = [str(len(item)) for item in val]
                        formatted_values.append(f"[{', '.join(lengths)}]")
                    else:
                        formatted_values.append(str(len(val)))
                values_str = ', '.join(formatted_values)
            else:
                values_str = ', '.join([str(val) for val in test_values])
        else:
            values_str = ''
        
        description_data.append({
            'Атрибут': field_name,
            'Тип данных': field_types.get(header, ''),
            'Пример': example,
            'minimum': min_val,
            'maximum': max_val,
            'maxLength': maxlen_val,
            'Значения для тестов': values_str,
            'Количество значений': len(test_values)
        })
    
    # Расчеты тестов
    total_combinations = 1
    test_values_counts = []
    
    for item in description_data:
        count = item['Количество значений']
        if count > 0:
            total_combinations *= count
            test_values_counts.append(count)
    
    pairwise_count = 0
    try:
        test_values_for_pairwise = []
        for header, example in zip(headers, examples):
            field_type = field_types.get(header, '').lower()
            min_val = min_values.get(header, '')
            max_val = max_values.get(header, '')
            maxlen_val = maxlen_values.get(header, '')
            
            values = []
            example_values = [ex.strip() for ex in example.split(';')] if example and ';' in example else [example] if example else []
            
            if any(t in field_type for t in ['string', 'text']) and 'array' not in field_type:
                values.extend([ex for ex in example_values if ex])
                if maxlen_val:
                    try:
                        max_len = int(maxlen_val)
                        values.append('a')
                        if max_len > 1:
                            values.append('a' * max_len)
                    except ValueError:
                        pass
            
            elif any(t in field_type for t in ['number', 'integer', 'int', 'float', 'double', 'decimal']) and 'array' not in field_type:
                for ex in example_values:
                    try:
                        num = float(ex) if '.' in ex else int(ex)
                        values.append(str(num))
                    except (ValueError, TypeError):
                        pass
                if min_val:
                    try:
                        min_num = float(min_val) if '.' in min_val else int(min_val)
                        values.append(str(min_num))
                    except ValueError:
                        pass
                if max_val:
                    try:
                        max_num = float(max_val) if '.' in max_val else int(max_val)
                        values.append(str(max_num))
                    except ValueError:
                        pass
            
            elif any(t in field_type for t in ['boolean', 'bool']):
                for ex in example_values:
                    if ex.lower() in ['true', 'false', '1', '0', 'y', 'n']:
                        bool_val = 'true' if ex.lower() in ['true', '1', 'y'] else 'false'
                        if bool_val not in values:
                            values.append(bool_val)
                if 'true' not in values:
                    values.append('true')
                if 'false' not in values:
                    values.append('false')
            
            elif 'array' in field_type:
                if example:
                    try:
                        if example.startswith('[') and example.endswith(']'):
                            parsed = json.loads(example)
                            if isinstance(parsed, list):
                                values.append(parsed)
                        else:
                            if 'string' in field_type:
                                values.append(example_values)
                            else:
                                values.append(example_values)
                    except:
                        values.append(example_values)
                if 'string' in field_type and maxlen_val:
                    try:
                        max_len = int(maxlen_val)
                        values.append(['a'])
                        if max_len > 1:
                            values.append(['a' * max_len])
                    except ValueError:
                        pass
            
            if not values:
                values.append(get_default_value(field_type))
            
            test_values_for_pairwise.append(values)
        
        pairwise_count = 0
        for _ in AllPairs(test_values_for_pairwise):
            pairwise_count += 1
            
    except Exception as e:
        print(f"  Ошибка расчета попарных тестов: {e}")
        pairwise_count = min(total_combinations, 100)
    
    if total_combinations > 0:
        efficiency = (pairwise_count / total_combinations) * 100
    else:
        efficiency = 0
    
    desc_df = pd.DataFrame(description_data)
    excel_path = output_dir / f"{sheet_name}_description.xlsx"
    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        desc_df.to_excel(writer, sheet_name='Описание атрибутов', index=False)
        
        calculations_data = {
            'Параметр': [
                'Название листа',
                'Количество атрибутов',
                'Количество значений для тестирования',
                'Общее количество комбинаций (полный перебор)',
                'Количество попарных тестов (реальное)',
                'Эффективность попарного тестирования',
                'Сокращение количества тестов'
            ],
            'Значение': [
                sheet_name,
                len(description_data),
                ', '.join([str(x) for x in test_values_counts]),
                f"{total_combinations:,}".replace(',', ' '),
                pairwise_count,
                f"{efficiency:.2f}%",
                f"Сокращено в {total_combinations/pairwise_count:.1f} раз" if pairwise_count > 0 else "N/A"
            ],
            'Описание': [
                'Имя обрабатываемого листа',
                'Общее количество тестируемых атрибутов',
                'Количество тестовых значений для каждого атрибута',
                'Количество тестов при полном переборе всех комбинаций',
                'Фактическое количество тестов при попарном тестировании',
                'Процент покрытия от полного перебора',
                'Во сколько раз сократилось количество тестов'
            ]
        }
        
        calc_df = pd.DataFrame(calculations_data)
        calc_df.to_excel(writer, sheet_name='Расчеты тестов', index=False)
        
        explanations_data = {
            'Поле': [
                'Атрибут',
                'Тип данных',
                'Пример',
                'minimum',
                'maximum', 
                'maxLength',
                'Значения для тестов',
                'Количество значений'
            ],
            'Описание': [
                'Имя атрибута (без полного пути)',
                'Тип данных из исходной таблицы',
                'Пример значения из исходной таблицы',
                'Минимальное значение для числовых полей',
                'Максимальное значение для числовых полей',
                'Максимальная длина для строковых полей',
                'Значения, используемые для тестирования (для строк выводятся длины в символах)',
                'Количество различных значений для тестирования'
            ]
        }
        
        explanations_df = pd.DataFrame(explanations_data)
        explanations_df.to_excel(writer, sheet_name='Пояснения', index=False)
    
    print(f"  Создан файл с описанием атрибутов: {excel_path}")
    print(f"  Расчеты: {total_combinations} комбинаций -> {pairwise_count} попарных тестов ({efficiency:.2f}% эффективность)")
    return excel_path

def convert_xlsx_to_postman(file_path: Path):
    """Конвертирует XLSX в Postman JSON с Post-response скриптом для каждого листа"""
    output_dir = file_path.parent / file_path.stem
    output_dir.mkdir(exist_ok=True)
    
    xls = pd.ExcelFile(file_path)
    processed_sheets = []
    skipped_sheets = []
    
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, dtype=str)
        header_row = df.iloc[0].fillna('').str.strip().tolist()
        
        try:
            name_idx = header_row.index('Наименование атрибута')
        except ValueError:
            print(f"  Ошибка на листе '{sheet_name}':")
            print(f"    Не найден столбец 'Наименование атрибута'")
            print("  Убедитесь в правильности заполнений столбцов и полей.")
            print("  Пропущен лист")
            skipped_sheets.append(sheet_name)
            continue

        try:
            type_idx = header_row.index('Тип данных')
        except ValueError:
            print(f"  Ошибка на листе '{sheet_name}':")
            print(f"    Не найден столбец 'Тип данных'")
            print("  Убедитесь в правильности заполнений столбцов и полей.")
            print("  Пропущен лист")
            skipped_sheets.append(sheet_name)
            continue

        try:
            req_idx = next(i for i, header in enumerate(header_row) if header.startswith(('Обязательность', 'Обязательное', 'Required')))
        except StopIteration:
            print(f"  Ошибка на листе '{sheet_name}':")
            print(f"    Не найден столбец, начинающийся с 'Обязательность', 'Обязательное' или 'Required'")
            print("  Убедитесь в правильности заполнений столбцов и полей.")
            print("  Пропущен лист")
            skipped_sheets.append(sheet_name)
            continue

        try:
            ex_idx = header_row.index('Пример')
        except ValueError:
            print(f"  Ошибка на листе '{sheet_name}':")
            print(f"    Не найден столбец 'Пример'")
            print("  Убедитесь в правильности заполнений столбцов и полей.")
            print("  Пропущен лист")
            skipped_sheets.append(sheet_name)
            continue

        try:
            edto_idx = header_row.index('Путь в eDTO')
        except ValueError:
            print(f"  Предупреждение на листе '{sheet_name}':")
            print(f"    Не найден столбец 'Путь в eDTO'. Проверки minimum, maximum, maxLength и другие не будут выполняться.")
            edto_idx = None

        try:
            min_idx = header_row.index('minimum')
        except ValueError:
            min_idx = None

        try:
            max_idx = header_row.index('maximum')
        except ValueError:
            max_idx = None

        try:
            maxlen_idx = header_row.index('maxLength')
        except ValueError:
            maxlen_idx = None

        data = df.iloc[1:].values.tolist()
        headers = []
        examples = []
        required_fields = []
        field_types = {}
        edto_paths = {}
        min_values = {}
        max_values = {}
        maxlen_values = {}
        
        current_path = []
        for row in data:
            row = ['' if pd.isna(x) else str(x).strip() for x in row]
            if not any(row):
                continue
                
            level = 0
            while level < len(row) and row[level] == '':
                level += 1
            
            if level >= len(row):
                continue
                
            name = row[level]
            field_type = row[type_idx] if type_idx < len(row) else ''
            is_required = row[req_idx] if req_idx < len(row) else ''
            example = row[ex_idx] if ex_idx < len(row) else ''
            edto_path = row[edto_idx] if edto_idx is not None and edto_idx < len(row) else ''
            min_val = row[min_idx] if min_idx is not None and min_idx < len(row) else ''
            max_val = row[max_idx] if max_idx is not None and max_idx < len(row) else ''
            maxlen_val = row[maxlen_idx] if maxlen_idx is not None and maxlen_idx < len(row) else ''
            
            current_path = current_path[:level]
            current_path.append(name)
            
            if level == 0:
                full_path = name
            else:
                if '[0]' in name:
                    array_part = name.replace('[0]', '') + '[0]'
                    full_path = '.'.join(current_path[:-1]) + '.' + array_part
                else:
                    full_path = '.'.join(current_path)
            
            headers.append(full_path)
            examples.append(example)
            field_types[full_path] = field_type
            if edto_path:
                edto_paths[full_path] = edto_path
            if min_val:
                min_values[full_path] = min_val
            if max_val:
                max_values[full_path] = max_val
            if maxlen_val:
                maxlen_values[full_path] = maxlen_val
            
            if is_required == 'О':
                required_fields.append(full_path)
        
        create_attributes_description_excel(file_path, sheet_name, headers, examples, field_types, 
                                          min_values, max_values, maxlen_values)
        
        create_postman_json(file_path, sheet_name, output_dir, headers, examples, field_types, required_fields, 
                           edto_paths, min_values, max_values, maxlen_values)
        processed_sheets.append(sheet_name)
    
    return processed_sheets, skipped_sheets

def build_required_json_structure(headers: List[str], examples: List[str], 
                                 field_types: Dict[str, str], required_fields: List[str]) -> Dict[str, Any]:
    """Строит JSON структуру только для обязательных полей"""
    result = {}
    
    for header, example in zip(headers, examples):
        if not header or header not in required_fields:
            continue
        
        field_type = field_types.get(header, '').lower()
        send_type = field_type.split('/')[0] if '/' in field_type else field_type
        parts = header.split('.')
        current = result
        
        for i, part in enumerate(parts):
            is_last = i == len(parts) - 1
            is_array = part.endswith('[0]')
            
            if is_array:
                array_name = part[:-3]
                
                if isinstance(current, dict):
                    if array_name not in current:
                        current[array_name] = []
                    current = current[array_name]
                
                elif isinstance(current, list):
                    if not current:
                        current.append({})
                    current = current[-1]
                
                if is_last:
                    value = parse_value(example, send_type)
                    if value is not None and value != '':
                        if isinstance(value, list):
                            if 'array objects' in send_type or 'array[object]' in send_type:
                                for v in value:
                                    current.append(v if isinstance(v, dict) else {"value": v})
                            else:
                                current.extend(value)
                        else:
                            current.append(value)
                else:
                    if isinstance(current, list) and not current:
                        current.append({})
                    if isinstance(current, list):
                        current = current[-1]
                    if not isinstance(current, dict):
                        current = current[-1] = {}
            else:
                if isinstance(current, list):
                    if not current:
                        current.append({})
                    current = current[-1]
                
                if not isinstance(current, dict):
                    raise ValueError(f"Ожидался словарь, получен {type(current)} для пути {header}")
                
                if is_last:
                    value = parse_value(example, send_type)
                    if value is not None and value != '':
                        current[part] = value
                else:
                    next_part = parts[i + 1] if i + 1 < len(parts) else ''
                    is_next_array = next_part.endswith('[0]')
                    
                    if part not in current:
                        current[part] = [] if is_next_array else {}
                    
                    current = current[part]
    
    def fix_arrays(obj):
        if isinstance(obj, dict):
            for key, value in list(obj.items()):
                if key.endswith('[0]'):
                    array_name = key[:-3]
                    if array_name not in obj:
                        obj[array_name] = [value]
                    else:
                        obj[array_name].append(value)
                    del obj[key]
                elif isinstance(value, list):
                    for item in value:
                        fix_arrays(item)
                elif isinstance(value, dict):
                    fix_arrays(value)
        return obj
    
    return fix_arrays(result)

def build_json_structure(headers: List[str], examples: List[str], 
                        field_types: Dict[str, str]) -> Dict[str, Any]:
    """Строит JSON структуру из плоского списка всех полей с учетом вложенных массивов"""
    result = {}
    
    for header, example in zip(headers, examples):
        if not header:
            continue
        
        field_type = field_types.get(header, '').lower()
        send_type = field_type.split('/')[0] if '/' in field_type else field_type
        parts = header.split('.')
        current = result
        
        for i, part in enumerate(parts):
            is_last = i == len(parts) - 1
            is_array = part.endswith('[0]')
            
            if is_array:
                array_name = part[:-3]
                
                if isinstance(current, dict):
                    if array_name not in current:
                        current[array_name] = []
                    current = current[array_name]
                
                elif isinstance(current, list):
                    if not current:
                        current.append({})
                    current = current[-1]
                
                if is_last:
                    value = parse_value(example, send_type)
                    if value is not None and value != '':
                        if isinstance(value, list):
                            if 'array objects' in send_type or 'array[object]' in send_type:
                                for v in value:
                                    current.append(v if isinstance(v, dict) else {"value": v})
                            else:
                                current.extend(value)
                        else:
                            current.append(value)
                else:
                    if isinstance(current, list) and not current:
                        current.append({})
                    if isinstance(current, list):
                        current = current[-1]
                    if not isinstance(current, dict):
                        current = current[-1] = {}
            else:
                if isinstance(current, list):
                    if not current:
                        current.append({})
                    current = current[-1]
                
                if not isinstance(current, dict):
                    raise ValueError(f"Ожидался словарь, получен {type(current)} для пути {header}")
                
                if is_last:
                    value = parse_value(example, send_type)
                    if value is not None and value != '':
                        current[part] = value
                else:
                    next_part = parts[i + 1] if i + 1 < len(parts) else ''
                    is_next_array = next_part.endswith('[0]')
                    
                    if part not in current:
                        current[part] = [] if is_next_array else {}
                    
                    current = current[part]
    
    def fix_arrays(obj):
        if isinstance(obj, dict):
            for key, value in list(obj.items()):
                if key.endswith('[0]'):
                    array_name = key[:-3]
                    if array_name not in obj:
                        obj[array_name] = [value]
                    else:
                        obj[array_name].append(value)
                    del obj[key]
                elif isinstance(value, list):
                    for item in value:
                        fix_arrays(item)
                elif isinstance(value, dict):
                    fix_arrays(value)
        return obj
    
    return fix_arrays(result)

def generate_test_values(headers: List[str], examples: List[str], field_types: Dict[str, str], 
                        min_values: Dict[str, str], max_values: Dict[str, str], maxlen_values: Dict[str, str]) -> List[List[Any]]:
    """Генерирует значения для попарного тестирования для каждого поля, используя тип до слэша."""
    test_values = []
    for header, example in zip(headers, examples):
        field_type = field_types.get(header, '').lower()
        send_type = field_type.split('/')[0] if '/' in field_type else field_type
        min_val = min_values.get(header, '')
        max_val = max_values.get(header, '')
        maxlen_val = maxlen_values.get(header, '')
        
        values = []
        example_values = [ex.strip() for ex in example.split(';')] if example and ';' in example else [example] if example else []
        
        if any(t in send_type for t in ['string', 'text']) and 'array' not in send_type:
            values.extend([ex for ex in example_values if ex])
            if maxlen_val:
                try:
                    max_len = int(maxlen_val)
                    min_string = ''.join(random.choices(string.ascii_letters, k=1))
                    values.append(min_string)
                    if max_len > 1:
                        max_string = ''.join(random.choices(string.ascii_letters, k=max_len))
                        values.append(max_string)
                except ValueError:
                    pass
        
        elif any(t in send_type for t in ['number', 'integer', 'int', 'float', 'double', 'decimal']) and 'array' not in send_type:
            for ex in example_values:
                try:
                    num = float(ex) if '.' in ex else int(ex)
                    values.append(str(num))
                except (ValueError, TypeError):
                    pass
            if min_val:
                try:
                    min_num = float(min_val) if '.' in min_val else int(min_val)
                    values.append(str(min_num))
                except ValueError:
                    pass
            if max_val:
                try:
                    max_num = float(max_val) if '.' in max_val else int(max_val)
                    values.append(str(max_num))
                except ValueError:
                    pass
        
        elif any(t in send_type for t in ['boolean', 'bool']):
            for ex in example_values:
                if ex.lower() in ['true', 'false', '1', '0', 'y', 'n']:
                    bool_val = 'true' if ex.lower() in ['true', '1', 'y'] else 'false'
                    if bool_val not in values:
                        values.append(bool_val)
            if 'true' not in values:
                values.append('true')
            if 'false' not in values:
                values.append('false')
        
        elif 'array' in send_type:
            if example:
                try:
                    if example.startswith('[') and example.endswith(']'):
                        parsed = json.loads(example)
                        if isinstance(parsed, list):
                            values.append(parsed)
                    else:
                        if 'string' in send_type:
                            values.append(example_values)
                        else:
                            values.append(example_values)
                except:
                    values.append(example_values)
            if 'string' in send_type and maxlen_val:
                try:
                    max_len = int(maxlen_val)
                    values.append(['a'])
                    if max_len > 1:
                        values.append(['a' * max_len])
                except ValueError:
                    pass
        
        if not values:
            values.append(get_default_value(send_type))
            print(f"  Для поля {header} использовано значение по умолчанию: {values[0]}")
        
        test_values.append([header, values])
    
    return test_values

def create_postman_json(file_path: Path, sheet_name: str, output_dir: Path, headers: List[str], 
                       examples: List[str], field_types: Dict[str, str], required_fields: List[str],
                       edto_paths: Dict[str, str], min_values: Dict[str, str], max_values: Dict[str, str],
                       maxlen_values: Dict[str, str]) -> None:
    """Создает JSON файл для импорта в Postman с запросами: только обязательные поля, все поля и попарные тесты.
    Также создает Excel-файл с попарными тестами."""
    required_json_structure = build_required_json_structure(headers, examples, field_types, required_fields)
    all_json_structure = build_json_structure(headers, examples, field_types)
    
    required_test_script = generate_post_response_script(headers, examples, field_types, required_fields, 
                                                        edto_paths, min_values, max_values, maxlen_values)
    all_test_script = generate_post_response_script(headers, examples, field_types, headers, 
                                                   edto_paths, min_values, max_values, maxlen_values)
    
    test_values = generate_test_values(headers, examples, field_types, min_values, max_values, maxlen_values)
    
    pairwise_requests = []
    pairwise_test_cases = []
    
    for i, pairs in enumerate(AllPairs([values for _, values in test_values])):
        test_case_values = {test_values[j][0]: value for j, value in enumerate(pairs)}
        
        test_examples = []
        test_case_row = {}
        for h in headers:
            if h in test_case_values:
                test_examples.append(test_case_values[h] if isinstance(test_case_values[h], str) else str(test_case_values[h]))
                test_case_row[h] = test_case_values[h]
            else:
                test_examples.append(examples[headers.index(h)])
                test_case_row[h] = examples[headers.index(h)]
        
        test_json_structure = build_json_structure(headers, test_examples, field_types)
        
        test_script = generate_post_response_script(headers, test_examples, field_types, headers, 
                                                  edto_paths, min_values, max_values, maxlen_values)
        
        pairwise_requests.append({
            "name": f"Test {i + 1}",
            "event": [
                {
                    "listen": "test",
                    "script": {
                        "type": "text/javascript",
                        "exec": test_script.splitlines()
                    }
                }
            ],
            "request": {
                "method": "POST",
                "header": [
                    {
                        "key": "Content-Type",
                        "value": "application/json"
                    }
                ],
                "body": {
                    "mode": "raw",
                    "raw": json.dumps(test_json_structure, ensure_ascii=False, indent=2)
                },
                "url": {
                    "raw": "{{base_url}}",
                    "host": ["{{base_url}}"]
                }
            },
            "response": []
        })
        
        pairwise_test_cases.append(test_case_row)
    
    pairwise_df = pd.DataFrame(pairwise_test_cases, index=[f"Test {i+1}" for i in range(len(pairwise_test_cases))])
    
    excel_path = output_dir / f"{sheet_name}_pairwise_tests.xlsx"
    pairwise_df.to_excel(excel_path, index=True, engine='openpyxl')
    print(f"  Сохранен Excel-файл с попарными тестами: {excel_path}")
    
    postman_data = {
        "info": {
            "name": f"API MAPPING - {file_path.stem} - {sheet_name}",
            "description": f"Автоматически сгенерированная коллекция из {file_path.name}, лист {sheet_name}",
            "schema": "https://schema.getpostman.com/json/collection/v2.1.0/collection.json"
        },
        "variable": [
            {
                "key": "base_url",
                "value": "https://your-api-domain.com",
                "type": "string",
                "description": "Базовый URL API"
            }
        ],
        "item": [
            {
                "name": f"Обязательные поля",
                "event": [
                    {
                        "listen": "test",
                        "script": {
                            "type": "text/javascript",
                            "exec": required_test_script.splitlines()
                        }
                    }
                ],
                "request": {
                    "method": "POST",
                    "header": [
                        {
                            "key": "Content-Type",
                            "value": "application/json"
                        }
                    ],
                    "body": {
                        "mode": "raw",
                        "raw": json.dumps(required_json_structure, ensure_ascii=False, indent=2)
                    },
                    "url": {
                        "raw": "{{base_url}}",
                        "host": ["{{base_url}}"]
                    }
                },
                "response": []
            },
            {
                "name": f"Все поля",
                "event": [
                    {
                        "listen": "test",
                        "script": {
                            "type": "text/javascript",
                            "exec": all_test_script.splitlines()
                        }
                    }
                ],
                "request": {
                    "method": "POST",
                    "header": [
                        {
                            "key": "Content-Type",
                            "value": "application/json"
                        }
                    ],
                    "body": {
                        "mode": "raw",
                        "raw": json.dumps(all_json_structure, ensure_ascii=False, indent=2)
                    },
                    "url": {
                        "raw": "{{base_url}}",
                        "host": ["{{base_url}}"]
                    }
                },
                "response": []
            }
        ] + pairwise_requests
    }
    
    postman_path = output_dir / f"{sheet_name}_postman.json"
    with open(postman_path, 'w', encoding='utf-8') as f:
        json.dump(postman_data, f, ensure_ascii=False, indent=2)
    print(f"  Сгенерировано {len(pairwise_requests)} попарных тестов для листа {sheet_name}")

def generate_post_response_script(headers: List[str], examples: List[str], 
                                field_types: Dict[str, str], fields_to_check: List[str],
                                edto_paths: Dict[str, str], min_values: Dict[str, str],
                                max_values: Dict[str, str], maxlen_values: Dict[str, str]) -> str:
    """Генерирует Post-response скрипт для Postman для указанных полей на русском языке, используя тип после слэша для проверок"""
    script = [
        "if (pm.response.code === 200) {",
        "    pm.test('Код ответа 200', function () {",
        "        pm.response.to.have.status(200);",
        "    });",
        "",
        "    pm.test('Тело ответа содержит валидный JSON', function () {",
        "        pm.response.to.be.json;",
        "        pm.response.to.not.be.error;",
        "    });",
        "",
        "    const response = pm.response.json();",
        ""
    ]

    for field in fields_to_check:
        if field not in headers:
            continue
        edto_path = edto_paths.get(field, '')
        if not edto_path:
            continue
            
        test_field_name = edto_path.split('.')[-1].replace('[0]', '') if edto_path else field.split('.')[-1]
        
        field_type = field_types.get(field, '').lower()
        # Используем тип после слэша для проверок, если он есть
        check_type = field_type.split('/')[1] if '/' in field_type else field_type
        # Тип до слэша используется для отправки данных
        send_type = field_type.split('/')[0] if '/' in field_type else field_type
        
        example = examples[headers.index(field)] if field in headers else ''
        min_val = min_values.get(field, '')
        max_val = max_values.get(field, '')
        maxlen_val = maxlen_values.get(field, '')
        
        script.append(f"    pm.test('Поле {test_field_name} присутствует', function () {{")
        script.append(f"        pm.expect(response).to.have.nested.property('{edto_path}');")
        
        script.append(f"        if (pm.expect(response).to.have.nested.property('{edto_path}')) {{")
        
        # Проверка типа данных - используем check_type (тип после слэша)
        if 'string' in check_type and 'array' not in check_type:
            script.append(f"            pm.test('Поле {test_field_name} является строкой', function () {{")
            script.append(f"                pm.expect(response.{edto_path}).to.be.a('string');")
            script.append("            });")
            
        elif any(t in check_type for t in ['number', 'integer', 'int', 'float', 'double', 'decimal']) and 'array' not in check_type:
            script.append(f"            pm.test('Поле {test_field_name} является числом', function () {{")
            script.append(f"                pm.expect(response.{edto_path}).to.be.a('number');")
            script.append("            });")
            
        elif any(t in check_type for t in ['boolean', 'bool']):
            script.append(f"            pm.test('Поле {test_field_name} является булевым значением', function () {{")
            script.append(f"                pm.expect(response.{edto_path}).to.be.a('boolean');")
            script.append("            });")
            
        elif 'object' in check_type and 'array' not in check_type:
            script.append(f"            pm.test('Поле {test_field_name} является объектом', function () {{")
            script.append(f"                pm.expect(response.{edto_path}).to.be.an('object');")
            script.append("            });")
            
        elif 'array' in check_type:
            script.append(f"            pm.test('Поле {test_field_name} является массивом', function () {{")
            script.append(f"                pm.expect(response.{edto_path}).to.be.an('array');")
            script.append("            });")
            
            # Для массивов с определенным типом элементов добавляем дополнительные проверки
            if 'string' in check_type:
                script.append(f"            pm.test('Элементы массива {test_field_name} являются строками', function () {{")
                script.append(f"                if (response.{edto_path}.length > 0) {{")
                script.append(f"                    pm.expect(response.{edto_path}[0]).to.be.a('string');")
                script.append("                }")
                script.append("            });")
                
            elif any(t in check_type for t in ['number', 'integer', 'int', 'float', 'double', 'decimal']):
                script.append(f"            pm.test('Элементы массива {test_field_name} являются числами', function () {{")
                script.append(f"                if (response.{edto_path}.length > 0) {{")
                script.append(f"                    pm.expect(response.{edto_path}[0]).to.be.a('number');")
                script.append("                }")
                script.append("            });")

        # Проверка значения - используем пример и преобразуем его в соответствии с check_type
        if example and example != '':
            try:
                # Для проверки значения используем тип проверки (check_type)
                parsed_example = parse_value(example, check_type)
                if isinstance(parsed_example, (dict, list)):
                    script.append(f"            pm.test('Поле {test_field_name} имеет ожидаемую структуру', function () {{")
                    script.append(f"                pm.expect(response.{edto_path}).to.deep.equal({json.dumps(parsed_example, ensure_ascii=False)});")
                    script.append("            });")
                else:
                    # Для массивов проверяем, что значение соответствует ожидаемому
                    if 'array' in check_type:
                        script.append(f"            pm.test('Поле {test_field_name} имеет ожидаемое значение', function () {{")
                        script.append(f"                pm.expect(response.{edto_path}).to.deep.equal({json.dumps(parsed_example, ensure_ascii=False)});")
                        script.append("            });")
                    else:
                        script.append(f"            pm.test('Поле {test_field_name} имеет ожидаемое значение', function () {{")
                        script.append(f"                pm.expect(response.{edto_path}).to.equal({json.dumps(parsed_example, ensure_ascii=False)});")
                        script.append("            });")
            except Exception as e:
                print(f"  Ошибка при обработке примера для поля {field}: {e}")
                # Если не удалось распарсить пример, используем строковое представление
                script.append(f"            pm.test('Поле {test_field_name} имеет ожидаемое значение', function () {{")
                script.append(f"                pm.expect(response.{edto_path}.toString()).to.equal('{example}');")
                script.append("            });")
        
        script.append("        }")
        script.append("    });")

    script.append("}")

    return "\n".join(script)

def parse_value(value: str, field_type: str) -> Any:
    """Парсит значение в соответствии с типом данных"""
    if isinstance(value, list):
        return value
    
    if not value or value == '' or value.lower() == 'null':
        return get_default_value(field_type)
    
    try:
        if isinstance(value, str):
            value = value.strip()
            if value.lower() == 'null':
                return get_default_value(field_type)
        else:
            return value
            
        # Убираем часть после слэша для парсинга значения
        parse_type = field_type.lower().split('/')[0] if '/' in field_type.lower() else field_type.lower()
        
        if 'string' in parse_type and 'array' not in parse_type:
            if 'date-time' in parse_type:
                return value
            return value
        
        elif any(t in parse_type for t in ['number', 'integer', 'int', 'float', 'double', 'decimal']) and 'array' not in parse_type:
            if '.' in value:
                return float(value)
            else:
                return int(value)
        
        elif any(t in parse_type for t in ['boolean', 'bool']):
            return value.lower() in ['true', '1', 'y']
        
        elif 'object' in parse_type and 'array' not in parse_type:
            try:
                if value.startswith('{') and value.endswith('}'):
                    return json.loads(value)
                else:
                    return {"value": value}
            except:
                return {"value": value}
        
        elif 'array' in parse_type:
            try:
                if value.startswith('[') and value.endswith(']'):
                    parsed = json.loads(value)
                    if isinstance(parsed, list):
                        return parsed
                values = [x.strip() for x in value.split(';')]
                
                if 'array objects' in parse_type or 'array[object]' in parse_type:
                    return [{"value": v} for v in values]
                elif 'array strings' in parse_type or 'array[string]' in parse_type:
                    return values
                elif 'array numbers' in parse_type or 'array[number]' in parse_type:
                    return [float(v) if '.' in v else int(v) for v in values]
                else:
                    return values
            except:
                return [value]
        
        else:
            return value
            
    except (ValueError, TypeError) as e:
        print(f"  Ошибка парсинга значения '{value}' типа '{field_type}': {e}")
        return value

def get_default_value(field_type: str) -> Any:
    """Возвращает значение по умолчанию для типа данных до слэша"""
    if not field_type:
        return ""
    
    send_type = field_type.lower().split('/')[0] if '/' in field_type.lower() else field_type.lower()
    
    if any(t in send_type for t in ['number', 'integer', 'int', 'float', 'double', 'decimal']) and 'array' not in send_type:
        return 0
    elif any(t in send_type for t in ['boolean', 'bool']):
        return False
    elif 'object' in send_type and 'array' not in send_type:
        return {}
    elif 'array' in send_type:
        if 'array objects' in send_type or 'array[object]' in send_type:
            return [{}]
        elif any(t in send_type for t in ['array numbers', 'array[number]']):
            return [0]
        else:
            return [""]
    else:
        return ""

def process_directory():
    """Обрабатывает все XLSX файлы в директории, исключая временные файлы"""
    clear_console()
    print("=" * 50)
    print("  Начало обработки XLSX файлов  ".center(50, "="))
    print("=" * 50 + "\n")
    
    current_dir = Path.cwd()
    xlsx_files = [file for file in current_dir.glob("*.xlsx") if not file.name.startswith('~$')]
    
    if not xlsx_files:
        print("  Внимание: XLSX файлы не найдены в текущей директории  ".center(50))
        print("=" * 50)
        return
    
    for file in xlsx_files:
        print(f"* Обработка файла: {file.name}")
        try:
            processed_sheets, skipped_sheets = convert_xlsx_to_postman(file)
            if processed_sheets:
                print(f"  Успешно обработано листов: {len(processed_sheets)}")
                print(f"  Создано файлов в папке '{file.stem}':")
                for sheet_name in processed_sheets:
                    print(f"    {sheet_name}_postman.json (Postman коллекция)")
                    print(f"    {sheet_name}_pairwise_tests.xlsx (Попарные тесты)")
                    print(f"    {sheet_name}_description.xlsx (Описание атрибутов)")
            if skipped_sheets:
                print(f"  Пропущено листов: {len(skipped_sheets)}")
                for sheet_name in skipped_sheets:
                    print(f"    {sheet_name}")
            if not processed_sheets and not skipped_sheets:
                print("  Не обработано ни одного листа")
        except Exception as e:
            print(f"  Ошибка обработки {file.name}:")
            print(f"    {str(e)}")
            print("  Убедитесь в правильности заполнений столбцов и полей.")
            print("  Повторите запуск кода")
        print("-" * 50 + "\n")
    
    print("=" * 50)
    print("  Обработка завершена! Все файлы сохранены  ".center(50, "="))
    print("=" * 50)

if __name__ == "__main__":
    process_directory()
