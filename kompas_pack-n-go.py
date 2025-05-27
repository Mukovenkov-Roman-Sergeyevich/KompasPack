import os
import shutil
import win32com.client
import pythoncom
import argparse
from pathlib import Path
from datetime import datetime

kompas_api_module = None 
kompas_constants = None                 

processed_files = set()
all_found_files = set()

def initialize_kompas_api(kompas_progid="Kompas.Application.7"):
    global kompas_api_module, kompas_constants
    iKompasApp_dispatch = None

    try:
        iKompasApp_dispatch = win32com.client.GetActiveObject(kompas_progid)
        print(f"Подключились к запущенному Компас-3D: ({kompas_progid}).")

    except pythoncom.com_error:
        print(f"Компас 3D с ProgId '{kompas_progid}' не запущен. Пытаемся запустить...")

        try:
            iKompasApp_dispatch = win32com.client.Dispatch(kompas_progid)
            print(f"Запущен Компас 3D ({kompas_progid}).")

        except pythoncom.com_error as e_dispatch:
            if kompas_progid == "Kompas.Application.7":
                print("Не получилось запустить Kompas.Application.7, пытаюсь Kompas.Application.5...")
                return initialize_kompas_api("Kompas.Application.5")

            else:
                print(f"Не удалось запустить Компас-3D {kompas_progid}: {e_dispatch}")
                return None

    iKompasApp_dispatch.Application.Visible = False 
    kompas_api7_constants = '{75C9F5D0-B5B8-4526-8681-9903C567D2ED}' 
    kompas_api7 = '{69AC2981-37C0-4379-84FD-5DD2F3C0A520}' 

    try:
        print(f"Пробуем загрузить API модуль (GUID: {kompas_api7})...")
        _temp_interfaces_module = win32com.client.gencache.EnsureModule(kompas_api7, 0, 1, 0)

        if not _temp_interfaces_module:
            raise RuntimeError(f"EnsureModule не нашёл модуль для интерфейсов (GUID: {kompas_api7}).")

        essential_interfaces = ['IKompasDocument', 'IKompasDocument3D', 'IKompasDocument2D', 'IPart7', 'IParts7', 'IModelObject']
        missing_interfaces = [iface for iface in essential_interfaces if not hasattr(_temp_interfaces_module, iface)]

        if missing_interfaces:
            print(f"Интерфейс модуль ({_temp_interfaces_module.__name__}) загружен, но отсутствуют необходимые интерфейсы {', '.join(missing_interfaces)}")
            print(f"Доступные аттрибуты: {str(dir(_temp_interfaces_module))[:300]}...") 
            kompas_api_module = None

        else:
            kompas_api_module = _temp_interfaces_module
            print(f"Успешно запущен API модуль Компас-3D {kompas_api_module.__name__}")
            for iface in essential_interfaces: print(f"  {iface} найден: Да")

    except Exception as e_interfaces_module:
        print(f"Ошибка при загрузке модуля для интерфейсов Компас-3D (GUID: {kompas_api7}): {e_interfaces_module}")
        kompas_api_module = None

    _temp_constants_module_wrapper = None

    try:
        print(f"Пробуем загрузить модуль для констант (GUID: {kompas_api7_constants})...")
        _temp_constants_module_wrapper = win32com.client.gencache.EnsureModule(kompas_api7_constants, 0, 1, 0)

        if not _temp_constants_module_wrapper:
            raise RuntimeError(f"EnsureModule не нашёл модуль для констант (GUID: {kompas_api7_constants})")

        print(f"Загружен API модуль для констант: {_temp_constants_module_wrapper.__name__}")
        if hasattr(_temp_constants_module_wrapper, 'constants') and hasattr(_temp_constants_module_wrapper.constants, 'ksDocumentAssembly'):
            print("Константы в модуле найдены.")
            kompas_constants = _temp_constants_module_wrapper.constants

        else:
            print(f"Константы не найдены (GUID: {kompas_api7_constants}).")
            kompas_constants = None

    except Exception as e_constants_module:
        print(f"Ошибка при загрузке модуля констант (GUID: {kompas_api7_constants}): {e_constants_module}")
        kompas_constants = None

    if not kompas_api_module:
        print(f"ФАТАЛЬНАЯ ОШИБКА: Модуль для интерфейсов не загружен (GUID: {kompas_api7})")
        if iKompasApp_dispatch: iKompasApp_dispatch.Application.Quit(); 
        return None

    if not kompas_constants or not hasattr(kompas_constants, 'ksDocumentAssembly'):
        print("ФАТАЛЬНАЯ ОШИБКА: Необходимые константы не были найдены")
        if iKompasApp_dispatch: iKompasApp_dispatch.Application.Quit(); 
        return None

    print("Успешная инициализация модуля констант и интерфейсов")
    return iKompasApp_dispatch

def find_dependencies_recursive(file_path_str, kompas_app_dispatch_obj):
    global kompas_api_module, kompas_constants 
    file_path = Path(file_path_str).resolve()
    normalized_file_path_str = str(file_path).lower()

    if normalized_file_path_str in processed_files: 
        print(f"  Файл {file_path} уже обработан, пропускаем.")
        return

    processed_files.add(normalized_file_path_str)
    if not file_path.exists():
        print(f"  Предупреждение: Файл не найден '{file_path}', пропускаем.")
        return

    all_found_files.add(file_path) 
    print(f" Обрабатываем: {file_path}")
    doc_dispatch = None

    try:
        doc_dispatch = kompas_app_dispatch_obj.Application.Documents.Open(str(file_path), False, True)
        if not doc_dispatch: 
            print(f"  Не удаётся открыть файл'{file_path}'. Невозможно получить зависимости.")
            return

        doc_type = -1 
        try:
            doc_typed_general = kompas_api_module.IKompasDocument(doc_dispatch)
            doc_type = doc_typed_general.DocumentType

        except AttributeError as e_IKompasDoc:
            print(f"  Невозможно определить тип документа: {e_IKompasDoc}")

        if doc_type == kompas_constants.ksDocumentAssembly:
            print(f"  Тип: Сборка")

            try:
                doc3D_interface = kompas_api_module.IKompasDocument3D(doc_dispatch)
                if not doc3D_interface: 
                    print("  Ошибка: doc3D_interface = None"); return

                top_part_interface = doc3D_interface.TopPart 
                if not top_part_interface: 
                    print("  Ошибка: IKompasDocument3D.TopPart = None."); return

                components_collection_interface = top_part_interface.Parts 
                if not components_collection_interface or components_collection_interface.Count == 0 : 
                    print("  Предупреждение: IPart7.Parts или пустой, или неправильно определился")

                else:
                    print(f"  Найдено {components_collection_interface.Count} компонентов в этой сборке")

                    for i in range(components_collection_interface.Count):
                        comp_item_dispatch = components_collection_interface.Item(i) 
                        if not comp_item_dispatch: 
                            print(f"  Компонент Item({i}) = None. Пропуск."); continue

                        try:
                            comp_as_ipart7 = kompas_api_module.IPart7(comp_item_dispatch)
                            if not comp_as_ipart7: 
                                print(f"  Ошибка: Item({i}) при cast в IPart7 = None. Пропуск"); continue

                            comp_file_path_str = comp_as_ipart7.FileName 
                            if comp_file_path_str:
                                comp_path_obj = Path(comp_file_path_str)
                                if not comp_path_obj.is_absolute():
                                    resolved_comp_path = (file_path.parent / comp_path_obj).resolve()
                                    print(f"    Определили '{comp_path_obj}' как '{resolved_comp_path}'")

                                else:
                                    resolved_comp_path = comp_path_obj.resolve()
                                find_dependencies_recursive(str(resolved_comp_path), kompas_app_dispatch_obj)

                            else: 
                                print(f"  Предупреждение: Item({i}) не имеет FileName или оно пустое.")

                        except AttributeError as e_cast_ipart7: 
                            print(f"  Не смогли обратиться к свойствам Item({i}): {e_cast_ipart7}")

            except AttributeError as e_attr_typed_asm: 
                print(f"  Сборка не смогла быть определена: {e_attr_typed_asm} (используемый модуль: {kompas_api_module.__name__ if kompas_api_module else 'Н/Д'})")

        elif doc_type == kompas_constants.ksDocumentDrawing:
            print(f"  Тип: Чертёж")

            try:
                doc2D_interface = kompas_api_module.IKompasDocument2D(doc_dispatch)
                if not doc2D_interface: 
                    print("  Ошибка. Чертёж не смог быть определён"); return

                layout_sheets = doc2D_interface.LayoutSheets
                if layout_sheets and layout_sheets.Count > 0:
                    for i in range(layout_sheets.Count):
                        sheet = layout_sheets.Item(i)
                        views = sheet.Views
                        if views and views.Count > 0:
                            for j in range(views.Count):
                                view = views.Item(j)
                                if hasattr(view, 'AssociatedModelFileName') and view.AssociatedModelFileName:
                                    model_file_path_str = view.AssociatedModelFileName
                                    model_path_obj = Path(model_file_path_str)
                                    if not model_path_obj.is_absolute():
                                        resolved_model_path = (file_path.parent / model_path_obj).resolve()

                                    else:
                                        resolved_model_path = model_path_obj.resolve()

                                    print(f"  Модель чертежа: {resolved_model_path}")
                                    find_dependencies_recursive(str(resolved_model_path), kompas_app_dispatch_obj)

                        else: print(f"  Лист {i} не имеет видов.")
                else: 
                    print("  У чертежа нет листов, либо была ошибка при их получении.")

            except AttributeError as e_attr_typed_drw: 
                print(f"  Не определился чертёж {e_attr_typed_drw}")

        elif doc_type == kompas_constants.ksDocumentPart:
            print("  Тип: Компонент.")

        else: 
            print(f"  Тип: Неизвестный тип ({doc_type}). Пропуск.")

    except Exception as e_outer_unexpected: 
        print(f"Внешняя неожиданная ошибка при обработке {file_path}: {e_outer_unexpected}")
        import traceback
        traceback.print_exc()

    finally:
        if doc_dispatch:
            try: 
                doc_dispatch.Close(False) 
            except Exception as e_close: 
                print(f"  Ошибка при закрытии документа {file_path}: {e_close}")

def update_paths_in_packed_assemblies(packed_dir_path, kompas_app_dispatch_obj):
    global kompas_api_module, kompas_constants
    print(f"\n--- Обновление путей компонентов: {packed_dir_path} ---")
    assembly_files_to_update = [f for f in packed_dir_path.glob('*.a3d')]

    if not assembly_files_to_update:
        print("Не найдены сборочные файлы *.a3d для обновления путей.")
        return

    for asm_file_path in assembly_files_to_update:
        print(f"  Обрабатываем сборочный файл для обновления путей: {asm_file_path.name}")
        doc_dispatch = None
        try:
            doc_dispatch = kompas_app_dispatch_obj.Application.Documents.Open(str(asm_file_path), False, False) 
            if not doc_dispatch:
                print(f"    Предупреждение: не был определён {asm_file_path.name} для обновления путей."); continue
            
            doc_typed_general = kompas_api_module.IKompasDocument(doc_dispatch)
            if doc_typed_general.DocumentType != kompas_constants.ksDocumentAssembly:
                print(f"    Предупреждение: {asm_file_path.name} не имеет тип сборка. Пропуск."); doc_dispatch.Close(True); continue

            doc3D_interface = kompas_api_module.IKompasDocument3D(doc_dispatch)
            if not doc3D_interface:
                print(f"    Предупреждение:{asm_file_path.name} не имеет тип IKomapsDocument. Пропуск."); doc_dispatch.Close(True); continue

            top_part_interface = doc3D_interface.TopPart
            if not top_part_interface:
                print(f"    Предупреждение: TopPart = None для {asm_file_path.name}. Пропуск."); doc_dispatch.Close(True); continue
            
            components_collection = top_part_interface.Parts
            if not components_collection or components_collection.Count == 0:
                print(f"    Не найдено компонентов для обновления."); doc_dispatch.Close(True); continue

            changes_made_to_this_asm = False
            print(f"    Найдено {components_collection.Count} компонентов. Проверяем пути...")
            for i in range(components_collection.Count):
                comp_item_dispatch = components_collection.Item(i)
                if not comp_item_dispatch: continue
                try:
                    comp_as_ipart7 = kompas_api_module.IPart7(comp_item_dispatch)
                    if not comp_as_ipart7: continue
                    current_full_path_str = comp_as_ipart7.FileName
                    if not current_full_path_str: print(f"      Компонент {i} не имеет FileName. Пропуск."); continue
                    
                    base_filename = Path(current_full_path_str).name 
                    new_relative_path = f".\\{base_filename}" 

                    if current_full_path_str != new_relative_path and Path(current_full_path_str).name != new_relative_path :
                        print(f"      Обновляем путь для '{base_filename}': до '{current_full_path_str}', после '{new_relative_path}'")
                        comp_as_ipart7.FileName = new_relative_path
                        
                        model_object_iface_for_update = kompas_api_module.IModelObject(comp_as_ipart7)
                        model_object_iface_for_update.Update()
                        changes_made_to_this_asm = True
                except Exception as e_comp_update:
                    print(f"      Ошибка при обработке компонента {i} для обновления путей: {e_comp_update}")

            if changes_made_to_this_asm:
                print(f"    Сохраняю изменения в {asm_file_path.name}..."); doc_dispatch.Save(); print(f"    {asm_file_path.name} сохранён.")
            
            doc_dispatch.Close(False) 
        except Exception as e_asm_update:
            print(f"    Неизвестная ошибка при обработке {asm_file_path.name}: {e_asm_update}")
            if doc_dispatch: 
                try: 
                    doc_dispatch.Close(True)
                except: 
                    pass
    print("--- Обновление путей закончено ---")


def main():
    parser = argparse.ArgumentParser(description="Компас-3D Pack-n-Go скрипт. Используйте его для передачи файлов")
    parser.add_argument("main_file", help="Путь к основному Компас-3D файлу (.a3d, .cdw).")
    parser.add_argument("output_dir", help="Куда сохранить сборку и зависимости? (Вне зависимости от опции --zip). Выходной путь.")
    parser.add_argument("--zip", help="Путь zip файла при желании. Абсолютный или релативный", nargs='?', const='packed_kompas_files.zip', type=str)
    parser.add_argument("--kompas_version", help="Компас ProgID. По умолчанию: Kompas.Application.7", default="Kompas.Application.7")
    parser.add_argument("--no_path_update", help="Без обновления зависимых путей. Необходимо будет восстановить оригинальную структуру или обновить все пути.", action="store_true")

    args = parser.parse_args()
    main_file_path = Path(args.main_file).resolve()
    base_output_path = Path(args.output_dir).resolve()

    if not main_file_path.exists(): print(f"Ошибка: Основной файл '{main_file_path}' не найден"); return

    actual_output_path = base_output_path 
    if base_output_path.exists():
        if not base_output_path.is_dir(): print(f"Ошибка: выходной путь'{base_output_path}' уже существует и не является директорией."); return
        if any(base_output_path.iterdir()): 
            print(f"Предупреждение: выходной путь '{base_output_path}' не пустой.")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            unique_subdir_name = f"{main_file_path.stem}_packed_{timestamp}"
            actual_output_path = base_output_path / unique_subdir_name

    try:
        actual_output_path.mkdir(parents=True, exist_ok=True) 
        print(f"Файлы будут здесь: {actual_output_path}")
    except OSError as e: print(f"Ошибка при создании директории {actual_output_path}: {e}"); return

    kompas_app_instance = None
    try:
        pythoncom.CoInitialize()
        kompas_app_instance = initialize_kompas_api(args.kompas_version)
        if not kompas_app_instance: print("При инициализации API произошла критическая ошибка"); return
        
        global processed_files, all_found_files
        processed_files.clear(); all_found_files.clear()

        print(f"\nПоиск зависимостей: {main_file_path}")
        find_dependencies_recursive(str(main_file_path), kompas_app_instance)
        print("\n--- Поиск завершен ---")
        if not all_found_files: print("Не было найдено файлов "); return
            
        sorted_files_to_pack = sorted(list(all_found_files), key=lambda p: str(p).lower())
        print(f"\nНайдено {len(sorted_files_to_pack)} уникальных файлов:")
        for f_path in sorted_files_to_pack: print(f"  - {f_path}")

        print(f"\nКопирую {len(sorted_files_to_pack)} файлов в {actual_output_path}...")
        copied_count, failed_count = 0, 0
        for src_obj in sorted_files_to_pack:
            dst_file = actual_output_path / src_obj.name
            if dst_file.resolve() == src_obj.resolve(): print(f"  Пропуск из-за идентичного пути: {src_obj.name}"); copied_count+=1; continue
            try: shutil.copy2(str(src_obj), dst_file); copied_count+=1
            except Exception as e: print(f"  Ошибка при копировании {src_obj} в {dst_file}: {e}"); failed_count+=1
        if copied_count > 0 : print(f"  Успешно скопировано {copied_count} файлов.")
        
        print(f"\n--- Копирование завершено --- \nВсего скопировано: {copied_count}. Всего пропущено/ошибочно: {failed_count}.")

        if copied_count > 0 and failed_count == 0 and not args.no_path_update:
            update_paths_in_packed_assemblies(actual_output_path, kompas_app_instance)

        elif args.no_path_update:
            print("\nНе обновляем зависимости")
        elif failed_count > 0:
            print("\nИз-за ошибок копирования, пропущено обновление зависимостей")
        else: 
            print("\Не было перемещено ни одного файла, пропущено обновление зависимостей.")

        if args.zip:
            zip_arg_name = Path(args.zip)
            final_zip_name_str = zip_arg_name.name if zip_arg_name.suffix.lower() == '.zip' else zip_arg_name.name + '.zip'
            zip_target_parent_dir = base_output_path.parent if str(base_output_path.parent) != str(base_output_path) else Path.cwd() 
            zip_target_path = zip_target_parent_dir / final_zip_name_str
            print(f"\nАрхивируем содержание '{actual_output_path}' в '{zip_target_path}'...")
            try:
                archive_base_name = str(zip_target_path.with_suffix(''))
                actual_zip_filepath = shutil.make_archive(archive_base_name, 'zip', root_dir=actual_output_path)
                print(f"Успешно создан ZIP файл: {actual_zip_filepath}")
            except Exception as e: print(f"Ошибка при создании ZIP файла: {e}")
    except Exception as e_toplevel_unexpected:
        print(f"Неожиданная ошибка: {e_toplevel_unexpected}"); import traceback; traceback.print_exc()
    finally:
        if kompas_app_instance: print("Компас продолжит работу.")
        pythoncom.CoUninitialize(); print("\nСкрипт завершен.")

if __name__ == "__main__":
    main()