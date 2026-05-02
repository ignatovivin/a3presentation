from __future__ import annotations


DIAGNOSTIC_RULE_LABELS: dict[str, str] = {
    "overflow_risk": "Текст может не поместиться",
    "font_bounds": "Размер шрифта вне допустимых границ",
    "rendered_text_overflow": "Текст вышел за границы блока",
    "capacity_retry": "Повторная разбивка по емкости",
    "split_for_capacity": "Слайд разбит по емкости",
    "split_overflow_slot": "Переполненный блок вынесен отдельно",
    "layout_degradation": "Макет применен с ограничениями",
    "inventory_fallback": "Использован запасной макет",
    "direct_shape_binding": "Использованы именованные shapes",
    "missing_table_shape": "Таблица не отрисована",
    "missing_chart_shape": "График не отрисован",
    "missing_image_shape": "Изображение не отрисовано",
    "missing_required_editable_slot": "В макете нет нужного редактируемого слота",
    "unexpected_table_shape": "Лишняя таблица на слайде",
    "unexpected_chart_shape": "Лишний график на слайде",
    "table_overlay_text_overflow": "Текст таблицы вышел за границы",
    "two_column_overlap": "Колонки пересекаются",
    "image_text_overlap": "Текст пересекается с изображением",
    "chart_type_mismatch": "Тип графика не совпал",
    "chart_series_count_mismatch": "Количество рядов графика не совпало",
    "chart_title_font_mismatch": "Шрифт заголовка графика не совпал",
    "chart_subtitle_font_mismatch": "Шрифт подзаголовка графика не совпал",
    "chart_value_axis_number_format_mismatch": "Формат оси графика не совпал",
    "chart_secondary_value_axis_number_format_mismatch": "Формат второй оси графика не совпал",
    "combo_chart_structure_mismatch": "Структура комбинированного графика не совпала",
    "missing_secondary_value_axis": "Не отрисована вторая ось графика",
    "narrow_table_content": "Область таблицы слишком узкая",
    "narrow_chart_content": "Область графика слишком узкая",
    "narrow_image_content": "Область изображения слишком узкая",
    "narrow_table_footer": "Footer таблицы слишком узкий",
    "narrow_footer": "Footer слишком узкий",
    "narrow_text_footer": "Footer текста слишком узкий",
    "footer_left_misalignment": "Footer смещен",
    "body_left_misalignment": "Текстовый блок смещен",
    "body_margin_mismatch": "Поля текстового блока не совпали",
    "narrow_image_panel": "Панель изображения слишком узкая",
    "image_text_body_misalignment": "Текст рядом с изображением смещен",
    "underfilled_placeholder_fill": "Блок заполнен заметно слабее ожидаемого",
    "underfilled_auxiliary_placeholder_fill": "Дополнительный блок не заполнен",
    "underfilled_subtitle_placeholder_fill": "Подзаголовок не заполнен",
    "content_footer_overlap": "Контент пересекается с footer",
    "title_subtitle_overlap": "Заголовок пересекается с подзаголовком",
    "subtitle_body_overlap": "Подзаголовок пересекается с текстом",
    "title_body_overlap": "Заголовок пересекается с текстом",
    "title_subtitle_gap_drift": "Зазор заголовка и подзаголовка изменился",
    "subtitle_body_gap_drift": "Зазор подзаголовка и текста изменился",
    "title_body_gap_drift": "Зазор заголовка и текста изменился",
    "content_order_mismatch": "Порядок контента изменился",
    "continuation_balance": "Продолжения слайдов несбалансированы",
    "underfilled_continuation": "Слайд-продолжение недозаполнен",
    "overflow_continuation": "Слайд-продолжение переполнен",
    "continuation_font_delta": "Шрифт продолжения отличается",
    "continuation_order_mismatch": "Порядок контента в продолжениях изменился",
    "background_fill_color_mismatch": "Цвет фона не совпал",
    "text_color_mismatch": "Цвет текста не совпал",
    "shape_fill_color_mismatch": "Цвет заливки фигуры не совпал",
    "shape_line_color_mismatch": "Цвет линии фигуры не совпал",
    "style_target_missing": "Целевой элемент стиля не найден",
    "table_header_fill_color_mismatch": "Цвет шапки таблицы не совпал",
    "table_header_text_color_mismatch": "Цвет текста шапки таблицы не совпал",
    "table_cell_text_overflow": "Текст в ячейке таблицы вышел за границы",
    "chart_series_color_missing": "Цвет ряда графика не применен",
}


DIAGNOSTIC_RULE_ACTIONS: dict[str, str] = {
    "overflow_risk": "Система попробует разбить текст на дополнительные слайды или выбрать более емкий макет.",
    "font_bounds": "Сократите текст или выберите макет с большим текстовым блоком.",
    "rendered_text_overflow": "Система повторит разбиение по фактическому переполненному блоку.",
    "capacity_retry": "Проверьте результат после автоматической повторной разбивки.",
    "split_for_capacity": "Проверьте порядок слайдов-продолжений.",
    "split_overflow_slot": "Проверьте слайд, куда вынесен переполненный блок.",
    "layout_degradation": "Проверьте выбранный макет в панели структуры.",
    "inventory_fallback": "Лучше выбрать полноценный PowerPoint layout вместо fallback.",
    "direct_shape_binding": "Проверьте, что именованные shapes в шаблоне сохранены.",
    "missing_table_shape": "Проверьте table placeholder или выберите табличный макет.",
    "missing_chart_shape": "Проверьте chart placeholder или выберите графический макет.",
    "missing_image_shape": "Проверьте image placeholder или выберите макет с изображением.",
    "missing_required_editable_slot": "Выберите target с нужной ролью или исправьте placeholder в шаблоне.",
    "unexpected_table_shape": "Проверьте соответствие типа слайда выбранному макету.",
    "unexpected_chart_shape": "Проверьте соответствие типа слайда выбранному макету.",
    "table_overlay_text_overflow": "Сократите значения в таблице или разбейте таблицу на несколько слайдов.",
    "two_column_overlap": "Выберите макет с большим зазором между колонками.",
    "image_text_overlap": "Выберите макет с большим зазором между изображением и текстом.",
    "chart_type_mismatch": "Проверьте выбранный тип графика в структуре документа.",
    "chart_series_count_mismatch": "Проверьте выбранные ряды графика и скрытые серии.",
    "combo_chart_structure_mismatch": "Проверьте распределение рядов между столбцами и линией.",
    "missing_secondary_value_axis": "Проверьте mixed-unit combo chart и вторичную ось.",
    "background_fill_color_mismatch": "Проверьте фон layout в шаблоне.",
    "text_color_mismatch": "Проверьте цвет текста в placeholder style шаблона.",
    "shape_fill_color_mismatch": "Проверьте заливку shape style шаблона.",
    "shape_line_color_mismatch": "Проверьте линию shape style шаблона.",
    "style_target_missing": "Проверьте, что целевой placeholder не удален из шаблона.",
    "table_header_fill_color_mismatch": "Проверьте стиль шапки таблицы в шаблоне.",
    "table_header_text_color_mismatch": "Проверьте цвет текста шапки таблицы в шаблоне.",
    "table_cell_text_overflow": "Сократите текст ячейки, увеличьте высоту строк или разбейте таблицу.",
    "chart_series_color_missing": "Проверьте палитру графиков в design tokens шаблона.",
}


def diagnostic_rule_label(rule: str) -> str:
    label = DIAGNOSTIC_RULE_LABELS.get(rule)
    if label:
        return label
    return rule.replace("_", " ").capitalize()


def diagnostic_rule_action(rule: str) -> str:
    action = DIAGNOSTIC_RULE_ACTIONS.get(rule)
    if action:
        return action
    return "Проверьте этот слайд в сгенерированной презентации."
