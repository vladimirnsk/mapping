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
import re

def parse_example_pairs(raw: str) -> List[tuple]:
    """Возвращает список кортежей (send, expected) для строки примера, разделённой через ';'.
    Примеры: 'ivan(text1);volodya(text2)' -> [('ivan','text1'), ('volodya','text2')]
    Если ожидаемого значения нет, second == None.
    """
    if not raw or raw == '' or raw is None:
        return []
    parts = [p.strip() for p in str(raw).split(';') if p.strip()]
    res = []
    for p in parts:
        m = re.match(r"^(?P<send>[^()]+?)(?:\((?P<check>.*)\))?$", p)
        if m:
            send = m.group('send').strip()
            check = m.group('check')
            if check is not None:
                check = check.strip()
            res.append((send, check))
        else:
            res.append((p, None))
    return res

def get_send_value(example: str) -> str:
    """Возвращает значение для отправки (часть до скобок) для данного примера.
    Если пример — JSON-массив ([...] ), возвращает исходный пример.
    Для нескольких значений через ';' возвращает первый send.
    """
    if not example or example == '':
        return example
    s = str(example).strip()
    if s.startswith('[') and s.endswith(']'):
        return s
    pairs = parse_example_pairs(s)
    if pairs:
        return pairs[0][0]
    # запасной вариант: если есть несколько значений, разделенных символом «;», вернуть первое
    if ';' in s:
        return s.split(';')[0].strip()
    return s

def get_expected_for_field(header: str, example_given: str, check_type: str, example_pairs_map: Dict[str, List[tuple]]) -> Any:
    """Определяет ожидаемое значение (для проверки ответа) для поля.
    Логика:
    - если в example_given есть скобки -> использовать их
    - иначе, если example_given — отправленное значение, искать соответствие в example_pairs_map[header]
    - иначе, попробовать распарсить example_given как значение через parse_value
    """
    # 1) попробуем распарсить прямо из example_given
    if example_given:
        parts = parse_example_pairs(example_given)
        if parts and parts[0][1] is not None:
            raw_expected = parts[0][1]
            return parse_value(raw_expected, check_type)

    # 2) если передано только отправленное значение — ищем в карте пар
    if example_pairs_map and header in example_pairs_map:
        pairs = example_pairs_map[header]
        # example_given может быть send-значением — ищем совпадение
        for send, expected in pairs:
            if send == example_given and expected is not None:
                return parse_value(expected, check_type)

    # 3) fallback — пытаемся распарсить сам example_given
    try:
        return parse_value(example_given, check_type)
    except Exception:
        return example_given

def clear_console():
    """Очищает консоль в зависимости от операционной системы"""
    os.system('cls' if os.name == 'nt' else 'clear')

def validate_pairwise_coverage(headers: List[str], test_values: List[List[Any]]) -> Dict[str, Any]:
    """
    Проверяет корректность покрытия попарного тестирования
    Возвращает статистику покрытия
    """
    parameters = [values for _, values in test_values]
    
    # Расчет общего количества пар
    total_possible_pairs = 0
    pair_combinations = []
    
    for i in range(len(parameters)):
        for j in range(i + 1, len(parameters)):
            pairs_count = len(parameters[i]) * len(parameters[j])
            total_possible_pairs += pairs_count
            pair_combinations.append({
                'param1': headers[i],
                'param2': headers[j], 
                'values1': len(parameters[i]),
                'values2': len(parameters[j]),
                'possible_pairs': pairs_count
            })
    
    # Генерация тестов и сбор покрытых пар
    covered_pairs = set()
    generated_tests = []
    
    for test_case in AllPairs(parameters):
        generated_tests.append(test_case)
        
        # Записываем все покрытые пары в этом тесте
        for i in range(len(test_case)):
            for j in range(i + 1, len(test_case)):
                # Преобразуем значения в строки для хеширования
                value1 = str(test_case[i]) if not isinstance(test_case[i], (str, int, float, bool)) else test_case[i]
                value2 = str(test_case[j]) if not isinstance(test_case[j], (str, int, float, bool)) else test_case[j]
                
                pair_key = (i, j, value1, value2)
                covered_pairs.add(pair_key)
    
    # Анализ покрытия по каждой паре параметров
    coverage_by_pair = []
    for pair_info in pair_combinations:
        i = headers.index(pair_info['param1'])
        j = headers.index(pair_info['param2'])
        
        covered_count = 0
        for pair in covered_pairs:
            if pair[0] == i and pair[1] == j:
                covered_count += 1
        
        coverage_percent = (covered_count / pair_info['possible_pairs']) * 100 if pair_info['possible_pairs'] > 0 else 100
        coverage_by_pair.append({
            **pair_info,
            'covered_pairs': covered_count,
            'coverage_percent': coverage_percent
        })
    
    return {
        'total_parameters': len(parameters),
        'total_possible_pairs': total_possible_pairs,
        'total_covered_pairs': len(covered_pairs),
        'total_tests': len(generated_tests),
        'overall_coverage': (len(covered_pairs) / total_possible_pairs) * 100 if total_possible_pairs > 0 else 100,
        'coverage_by_pair': coverage_by_pair,
        'generated_tests': generated_tests
    }

def create_coverage_report(coverage_stats: Dict[str, Any], output_path: Path, sheet_name: str):
    """Создает детальный отчет о покрытии в Excel"""
    
    # Общая статистика
    summary_data = {
        'Метрика': [
            'Всего параметров',
            'Всего возможных пар', 
            'Покрыто пар',
            'Сгенерировано тестов',
            'Общее покрытие пар',
            'Эффективность тестирования'
        ],
        'Значение': [
            coverage_stats['total_parameters'],
            coverage_stats['total_possible_pairs'],
            coverage_stats['total_covered_pairs'],
            coverage_stats['total_tests'],
            f"{coverage_stats['overall_coverage']:.6f}%",
            f"Сокращено в {coverage_stats['total_possible_pairs']/coverage_stats['total_tests']:.1f} раз" if coverage_stats['total_tests'] > 0 else "N/A"
        ],
        'Описание': [
            'Количество тестируемых параметров (полей)',
            'Общее количество возможных пар значений',
            'Фактически покрыто пар значений',
            'Количество сгенерированных тест-кейсов',
            'Процент покрытия всех возможных пар',
            'Во сколько раз сокращено количество тестов по сравнению с полным перебором'
        ]
    }
    
    # Детали по парам
    pair_coverage_data = []
    for pair in coverage_stats['coverage_by_pair']:
        pair_coverage_data.append({
            'Параметр 1': pair['param1'],
            'Значений в параметре 1': pair['values1'],
            'Параметр 2': pair['param2'],
            'Значений в параметре 2': pair['values2'], 
            'Возможных пар': pair['possible_pairs'],
            'Покрыто пар': pair['covered_pairs'],
            'Покрытие %': f"{pair['coverage_percent']:.2f}%",
            'Статус': 'Полное покрытие' if pair['coverage_percent'] == 100 else 'Неполное покрытие'
        })
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Сводка
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Сводка покрытия', index=False)
        
        # Детальное покрытие
        coverage_df = pd.DataFrame(pair_coverage_data)
        coverage_df.to_excel(writer, sheet_name='Покрытие по парам', index=False)

def create_attributes_description_excel(file_path: Path, sheet_name: str, headers: List[str], examples: List[str], 
                                      field_types: Dict[str, str], min_values: Dict[str, str], 
                                      max_values: Dict[str, str], maxlen_values: Dict[str, str]) -> Path:
    """Создает отдельный Excel файл с описанием атрибутов для конкретного листа"""
    output_dir = file_path.parent / file_path.stem
    output_dir.mkdir(exist_ok=True)
    
    description_data = []
    
    for header, example in zip(headers, examples):
        field_type = field_types.get(header, '').lower()
        min_val = min_values.get(header, '')
        max_val = max_values.get(header, '')
        maxlen_val = maxlen_values.get(header, '')

        # Получаем только имя поля (последнюю часть пути)
        field_name = header.split('.')[-1].replace('[0]', '')

        # Формируем значения для тестирования
        test_values = []

        # Разделяем значения из столбца "Пример" по символу ";" и поддерживаем синтаксис send(expected)
        pairs = parse_example_pairs(example)
        example_values = [send for send, _ in pairs] if pairs else ([ex.strip() for ex in example.split(';')] if example and ';' in example else [example] if example else [])
        has_expected_in_parentheses = any(check for _, check in pairs)
        
        # Для строковых типов
        if any(t in field_type for t in ['string', 'text']) and 'array' not in field_type:
            # Если примеры содержат только Y/N, используем ровно 'Y' и 'N'
            non_empty = [ex for ex in example_values if ex and str(ex).strip()]
            if non_empty and all(str(ex).strip().lower() in ['y', 'n'] for ex in non_empty):
                if 'Y' not in test_values:
                    test_values.append('Y')
                if 'N' not in test_values:
                    test_values.append('N')
            else:
                # Добавляем все значения из примера как отдельные
                for ex in example_values:
                    if ex and ex not in test_values:
                        test_values.append(ex)
                # Если в "Примере" есть ожидаемые значения в скобках, не добавляем min/max/maxLength-значения
                if (not has_expected_in_parentheses) and maxlen_val:
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
        
        # Для числовых типов
        elif any(t in field_type for t in ['number', 'integer', 'int', 'float', 'double', 'decimal']) and 'array' not in field_type:
            # Пробуем преобразовать примеры в числа
            for ex in example_values:
                try:
                    num = float(ex) if '.' in ex else int(ex)
                    if str(num) not in test_values:
                        test_values.append(str(num))
                except (ValueError, TypeError):
                    pass
            # Не добавляем min/max, если есть ожидаемые значения в скобках
            if (not has_expected_in_parentheses) and min_val and min_val not in test_values:
                try:
                    min_num = float(min_val) if '.' in min_val else int(min_val)
                    test_values.append(str(min_num))
                except ValueError:
                    pass
            if (not has_expected_in_parentheses) and max_val and max_val not in test_values:
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
            
            # Для массивов со скобками не добавляем дополнительные maxlen-элементы
            if (not has_expected_in_parentheses) and 'string' in field_type and maxlen_val:
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
    
    desc_df = pd.DataFrame(description_data)
    excel_path = output_dir / f"{sheet_name}_description.xlsx"
    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:

        # Сохранение описания атрибутов в Excel 
        desc_df.to_excel(writer, sheet_name='Описание атрибутов', index=False)
        
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
        
        startcol = 8 + 2
        explanations_df.to_excel(writer, sheet_name='Описание атрибутов', startcol=startcol, index=False)

    print(f"  Создан файл с описанием атрибутов: {excel_path}")
    return excel_path

def parse_edto_segments(edto_path: str) -> List[tuple[str, str | None, str | None]]:
    """Парсит путь eDTO в сегменты: (field, cond_key, cond_val)"""
    if not edto_path:
        return []
    parts = re.split(r'\.', edto_path)
    segments = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        match = re.match(r'([^\[]+)(?:\[([^]]+)\])?', part)
        if not match:
            segments.append((part, None, None))
            continue
        field = match.group(1).strip()
        cond_str = match.group(2)
        cond_key = None
        cond_val = None
        if cond_str:
            eq_pos = cond_str.find('=')
            if eq_pos != -1:
                key_str = cond_str[:eq_pos].strip()
                val_str = cond_str[eq_pos + 1:].strip()
                if len(val_str) >= 2 and val_str[0] in '"\'' and val_str[0] == val_str[-1]:
                    val_str = val_str[1:-1].strip()
                cond_key = key_str
                cond_val = val_str
        segments.append((field, cond_key, cond_val))
    return segments

def generate_navigation_code(script: List[str], segments: List[tuple[str, str | None, str | None]]) -> None:
    """Генерирует код навигации по сегментам в скрипт."""
    script.append("        let current = response;")
    for f, k, v in segments:
        script.append(f"        current = current.{f};")
        if k is not None:
            script.append(f"        if (Array.isArray(current)) {{")
            script.append(f"            current = current.find(item => item.{k} === {json.dumps(v)});")
            script.append("            pm.expect(current).to.not.be.undefined;")
            script.append("        }} else {{")
            script.append(f"            pm.expect(current.{k}).to.equal({json.dumps(v)});" )
            script.append("        }}")

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
        
        # Создаем файл описания атрибутов
        desc_excel_path = create_attributes_description_excel(file_path, sheet_name, headers, examples, field_types, 
                          min_values, max_values, maxlen_values)

        create_postman_json(file_path, sheet_name, output_dir, headers, examples, field_types, required_fields, 
                           edto_paths, min_values, max_values, maxlen_values, description_excel_path=desc_excel_path)
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
        pairs = parse_example_pairs(example)
        example_values = [send for send, _ in pairs] if pairs else ([ex.strip() for ex in example.split(';')] if example and ';' in example else [example] if example else [])
        has_expected_in_parentheses = any(check for _, check in pairs)
        
        if any(t in send_type for t in ['string', 'text']) and 'array' not in send_type:
            non_empty = [ex for ex in example_values if ex and str(ex).strip()]
            if non_empty and all(str(ex).strip().lower() in ['y', 'n'] for ex in non_empty):
                if 'Y' not in values:
                    values.append('Y')
                if 'N' not in values:
                    values.append('N')
            else:
                values.extend([ex for ex in example_values if ex])
                # don't add generated max/min strings when expected-in-parentheses present
                if (not has_expected_in_parentheses) and maxlen_val:
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
            # don't add maxlen-based array items when expected-in-parentheses present
            if (not has_expected_in_parentheses) and 'string' in send_type and maxlen_val:
                try:
                    max_len = int(maxlen_val)
                    values.append(['a'])
                    if max_len > 1:
                        values.append(['a' * max_len])
                except ValueError:
                    pass
        
        # Добавляем значение по умолчанию, если ничего не найдено
        if not values:
            values.append(get_default_value(send_type))
        
        test_values.append([header, values])
    
    return test_values

def create_postman_json(file_path: Path, sheet_name: str, output_dir: Path, headers: List[str], 
                       examples: List[str], field_types: Dict[str, str], required_fields: List[str],
                       edto_paths: Dict[str, str], min_values: Dict[str, str], max_values: Dict[str, str],
                       maxlen_values: Dict[str, str], description_excel_path: Path = None) -> None:
    """Создает JSON файл для импорта в Postman с запросами: только обязательные поля, все поля и попарные тесты.
    Также создает Excel-файл с попарными тестами."""
    # Используем отправляемые значения (без части в скобках) для формирования тел запросов
    send_examples = [get_send_value(ex) for ex in examples]

    required_json_structure = build_required_json_structure(headers, send_examples, field_types, required_fields)
    all_json_structure = build_json_structure(headers, send_examples, field_types)
    
    # Подготовим карту пар (send, expected) для каждого поля — понадобится для проверок
    example_pairs_map: Dict[str, List[tuple]] = {}
    for h, ex in zip(headers, examples):
        example_pairs_map[h] = parse_example_pairs(ex)

    required_test_script = generate_post_response_script(headers, examples, field_types, required_fields, 
                                                        edto_paths, min_values, max_values, maxlen_values, example_pairs_map)
    all_test_script = generate_post_response_script(headers, examples, field_types, headers, 
                                                   edto_paths, min_values, max_values, maxlen_values, example_pairs_map)
    
    test_values = generate_test_values(headers, examples, field_types, min_values, max_values, maxlen_values)
    
    # ПРОВЕРКА ПОКРЫТИЯ ALLPAIRS
    print(f" \n Проверка покрытия попарного тестирования для листа '{sheet_name}'...")
    coverage_stats = validate_pairwise_coverage(headers, test_values)
    
    print(f"  РЕЗУЛЬТАТЫ ПРОВЕРКИ ПОКРЫТИЯ:")
    print(f"    Всего параметров: {coverage_stats['total_parameters']}")
    print(f"    Всего возможных пар: {coverage_stats['total_possible_pairs']}")
    print(f"    Покрыто пар: {coverage_stats['total_covered_pairs']}")
    print(f"    Сгенерировано тестов: {coverage_stats['total_tests']}")
    print(f"    Общее покрытие пар: {coverage_stats['overall_coverage']:.2f}%")
    
    # Количество пар с неполным покрытием
    low_coverage_pairs = [p for p in coverage_stats['coverage_by_pair'] if p['coverage_percent'] < 100]
    total_pairs = coverage_stats['total_possible_pairs']
    if low_coverage_pairs:
        print(f"    Количество пар с неполным покрытием: {len(low_coverage_pairs)} из {total_pairs}")
    else:
        print(f"    Все пары имеют 100% покрытие!")

     # Создание отчета о покрытии
    coverage_report_path = output_dir / f"{sheet_name}_coverage_report.xlsx"
    create_coverage_report(coverage_stats, coverage_report_path, sheet_name)
    print(f"  Создан отчет о покрытии: {coverage_report_path}")
    
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

        # Для тела запроса используем отправляемые значения (без ожидаемой части в скобках)
        test_send_examples = [get_send_value(ex) for ex in test_examples]
        test_json_structure = build_json_structure(headers, test_send_examples, field_types)

        test_script = generate_post_response_script(headers, test_examples, field_types, headers, 
                          edto_paths, min_values, max_values, maxlen_values, example_pairs_map)

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
        
        # Для Excel с попарными тестами сохраняем именно отправляемые значения (без ожидаемой части в скобках)
        send_case_row = {h: get_send_value(test_case_row[h]) if test_case_row.get(h) is not None else '' for h in headers}
        pairwise_test_cases.append(send_case_row)
    
    pairwise_df = pd.DataFrame(pairwise_test_cases, index=[f"Test {i+1}" for i in range(len(pairwise_test_cases))])
    
    # Если передан путь к файлу описания атрибутов, добавляем лист с попарными тестами туда
    if description_excel_path:
        try:
            with pd.ExcelWriter(description_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                pairwise_df.to_excel(writer, sheet_name='Попарные тесты', index=True)
        except Exception as e:
            # fallback: сохраняем в отдельный файл
            excel_path = output_dir / f"{sheet_name}_pairwise_tests.xlsx"
            pairwise_df.to_excel(excel_path, index=True, engine='openpyxl')
            print(f"  Не удалось добавить лист в файл описания атрибутов ({e}), сохранено в: {excel_path}")
    else:
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

def generate_post_response_script(headers: List[str], examples: List[str], 
                                field_types: Dict[str, str], fields_to_check: List[str],
                                edto_paths: Dict[str, str], min_values: Dict[str, str],
                                max_values: Dict[str, str], maxlen_values: Dict[str, str],
                                example_pairs_map: Dict[str, List[tuple]]) -> str:
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
        
        test_field_name = field.split('.')[-1]
        field_type = field_types.get(field, '').lower()
        # Используем тип после слэша для проверок, если он есть
        check_type = field_type.split('/')[1] if '/' in field_type else field_type
        # Тип до слэша используется для отправки данных
        send_type = field_type.split('/')[0] if '/' in field_type else field_type

        example = examples[headers.index(field)] if field in headers else ''
        min_val = min_values.get(field, '')
        max_val = max_values.get(field, '')
        maxlen_val = maxlen_values.get(field, '')

        segments = parse_edto_segments(edto_path)
        simple_path = '.'.join([f for f, _, _ in segments])
        has_conditions = any(k is not None for _, k, _ in segments)

        expected_value = None
        try:
            expected_value = get_expected_for_field(field, example, check_type, example_pairs_map)
        except Exception as e:
            print(f"  Ошибка при определении ожидаемого значения для поля {field}: {e}")

        if has_conditions:
            # Дополнительные проверки для условий (количество = число скобок с условиями)
            for c_idx in range(len(segments)):
                _, cond_key, cond_val = segments[c_idx]
                if cond_key is None:
                    continue
                script.append(f"    pm.test('Поле {cond_key} имеет значение {cond_val}', function () {{")
                script.append("        let current = response;")
                # Предыдущие сегменты с фильтрами
                for s_idx in range(c_idx):
                    s_f, s_k, s_v = segments[s_idx]
                    script.append(f"        current = current.{s_f};")
                    if s_k is not None:
                        script.append(f"        if (Array.isArray(current)) {{")
                        script.append(f"            current = current.find(item => item.{s_k} === {json.dumps(s_v)});")
                        script.append("            pm.expect(current).to.not.be.undefined;")
                        script.append("        }} else {{")
                        script.append(f"            pm.expect(current.{s_k}).to.equal({json.dumps(s_v)});")
                        script.append("        }}")
                # Текущий сегмент
                s_f, s_k, s_v = segments[c_idx]
                script.append(f"        current = current.{s_f};")
                script.append(f"        if (Array.isArray(current)) {{")
                script.append(f"            const match = current.find(item => item.{s_k} === {json.dumps(s_v)});")
                script.append("            pm.expect(match).to.not.be.undefined;")
                script.append(f"            pm.expect(match.{s_k}).to.equal({json.dumps(s_v)});")
                script.append("        }} else {{")
                script.append(f"            pm.expect(current.{s_k}).to.equal({json.dumps(s_v)});")
                script.append("        }}")
                script.append("    }});")

            # Основная проверка присутствия
            script.append(f"    pm.test('Поле {test_field_name} присутствует', function () {{")
            generate_navigation_code(script, segments)
            script.append("        pm.expect(current).to.not.be.undefined;")
            script.append("    }});")

            # Проверка типа
            script.append(f"    pm.test('Поле {test_field_name} является корректного типа', function () {{")
            generate_navigation_code(script, segments)
            if 'string' in check_type and 'array' not in check_type:
                script.append("        pm.expect(current).to.be.a('string');")
            elif any(t in check_type for t in ['number', 'integer', 'int', 'float', 'double', 'decimal']) and 'array' not in check_type:
                script.append("        pm.expect(current).to.be.a('number');")
            elif any(t in check_type for t in ['boolean', 'bool']):
                script.append("        pm.expect(current).to.be.a('boolean');")
            elif 'object' in check_type and 'array' not in check_type:
                script.append("        pm.expect(current).to.be.an('object');")
            elif 'array' in check_type:
                script.append("        pm.expect(current).to.be.an('array');")
                if 'string' in check_type:
                    script.append("        if (current && current.length > 0) {")
                    script.append("            pm.expect(current[0]).to.be.a('string');")
                    script.append("        }")
                elif any(t in check_type for t in ['number', 'integer', 'int', 'float', 'double', 'decimal']):
                    script.append("        if (current && current.length > 0) {")
                    script.append("            pm.expect(current[0]).to.be.a('number');")
                    script.append("        }")
            script.append("    }});")

            # Проверка значения, если есть
            if expected_value is not None and expected_value != '':
                script.append(f"    pm.test('Поле {test_field_name} имеет ожидаемое значение', function () {{")
                generate_navigation_code(script, segments)
                if isinstance(expected_value, (dict, list)):
                    script.append(f"        pm.expect(current).to.deep.equal({json.dumps(expected_value, ensure_ascii=False)});")
                else:
                    val_repr = json.dumps(expected_value, ensure_ascii=False)
                    script.append(f"        pm.expect(current).to.equal({val_repr});")
                script.append("    }});")
        else:
            # Простой случай без условий
            script.append(f"    pm.test('Поле {test_field_name} присутствует', function () {{")
            script.append(f"        pm.expect(response).to.have.nested.property('{edto_path}');")
            script.append(f"        if (pm.expect(response).to.have.nested.property('{edto_path}')) {{")
            
            # Проверка типа данных - используем check_type (тип после слэша)
            if 'string' in check_type and 'array' not in check_type:
                script.append(f"            pm.test('Поле {test_field_name} является строкой', function () {{")
                script.append(f"                pm.expect(_.get(response, '{edto_path}')).to.be.a('string');")
                script.append("            });")
                
            elif any(t in check_type for t in ['number', 'integer', 'int', 'float', 'double', 'decimal']) and 'array' not in check_type:
                script.append(f"            pm.test('Поле {test_field_name} является числом', function () {{")
                script.append(f"                pm.expect(_.get(response, '{edto_path}')).to.be.a('number');")
                script.append("            });")
                
            elif any(t in check_type for t in ['boolean', 'bool']):
                script.append(f"            pm.test('Поле {test_field_name} является булевым значением', function () {{")
                script.append(f"                pm.expect(_.get(response, '{edto_path}')).to.be.a('boolean');")
                script.append("            });")
                
            elif 'object' in check_type and 'array' not in check_type:
                script.append(f"            pm.test('Поле {test_field_name} является объектом', function () {{")
                script.append(f"                pm.expect(_.get(response, '{edto_path}')).to.be.an('object');")
                script.append("            });")
                
            elif 'array' in check_type:
                script.append(f"            pm.test('Поле {test_field_name} является массивом', function () {{")
                script.append(f"                pm.expect(_.get(response, '{edto_path}')).to.be.an('array');")
                script.append("            });")
                
                # Для массивов с определенным типом элементов добавляем дополнительные проверки
                if 'string' in check_type:
                    script.append(f"            pm.test('Элементы массива {test_field_name} являются строками', function () {{")
                    script.append(f"                if (_.get(response, '{edto_path}') && _.get(response, '{edto_path}').length > 0) {{")
                    script.append(f"                    pm.expect(_.get(response, '{edto_path}')[0]).to.be.a('string');")
                    script.append("                }")
                    script.append("            });")
                    
                elif any(t in check_type for t in ['number', 'integer', 'int', 'float', 'double', 'decimal']):
                    script.append(f"            pm.test('Элементы массива {test_field_name} являются числами', function () {{")
                    script.append(f"                if (_.get(response, '{edto_path}') && _.get(response, '{edto_path}').length > 0) {{")
                    script.append(f"                    pm.expect(_.get(response, '{edto_path}')[0]).to.be.a('number');")
                    script.append("                }")
                    script.append("            });")

            # Проверка значения - используем ожидаемое значение из скобок (если есть)
            if expected_value is not None and expected_value != '':
                # все кроме структур и массивов
                if not isinstance(expected_value, (dict, list)):
                    script.append(f"            pm.test('Поле {test_field_name} имеет ожидаемое значение', function () {{")
                    # Для строк/чисел используем to.equal
                    val_repr = json.dumps(expected_value, ensure_ascii=False)
                    script.append(f"                pm.expect(_.get(response, '{edto_path}')).to.equal({val_repr});")
                    script.append("            });")
                else:
                    script.append(f"            pm.test('Поле {test_field_name} имеет ожидаемое значение', function () {{")
                    val_repr = json.dumps(expected_value, ensure_ascii=False)
                    script.append(f"                pm.expect(_.get(response, '{edto_path}')).to.deep.equal({val_repr});")
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
            # Поддержка нескольких чисел через ";" в одном поле, например '11;13'
            if ';' in value:
                parts = [p.strip() for p in value.split(';') if p.strip()]
                nums = []
                for p in parts:
                    try:
                        nums.append(float(p) if '.' in p else int(p))
                    except Exception:
                        nums.append(p)
                return nums
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
                print(f" \n Успешно обработано листов: {len(processed_sheets)}")
                print(f"  Создано файлов в папке '{file.stem}':")
                for sheet_name in processed_sheets:
                    print(f"    {sheet_name}_postman.json (Postman коллекция)")
                    print(f"    {sheet_name}_description.xlsx (Описание тестов)")
                    print(f"    {sheet_name}_coverage_report.xlsx (Отчет о покрытии)")
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