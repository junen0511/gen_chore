import json
import time
import os
from string import Template
import pandas as pd

# read file url
target_url = '/Users/junen/Downloads/YTO-INTL-EDI-Platform中英文文案20230315.xlsx'

# write options
time_new = time.strftime("%Y%m%d%H%M", time.localtime())
write_path = "/Users/junen/Downloads/{}".format('EDI_' + time_new)

# 判断是否存在当前地址
if not os.path.exists(write_path):
    os.makedirs(write_path)


def getFormTemplate(lang_labels={}):
    radio_group_str = " initialValue='true' options={['true', 'false',]}"
    value_enum_str = " valueEnum={{1: 'dic_01',2: 'dic_02'}}"
    start_tag_str = '<ProForm$type'
    end_tag_str = ' />'
    form_item_str = " colProps={{ md: 12, xl: 8 }} name='$name' label={intl.formatMessage({id: '$id',defaultMessage:'$name'})}"
    if lang_labels['type'] == 'Select':
        form_item_str = form_item_str + value_enum_str
    elif lang_labels['type'] == 'Radio':
        form_item_str = '.Group ' + form_item_str + radio_group_str

    form_item_str = start_tag_str + form_item_str + end_tag_str
    return Template(form_item_str).substitute(lang_labels)


def getTableTemplate(search_type='', lang_values={}):
    start_tag_str = '{'
    end_tag_str = '},'
    column_str = " title: <FormattedMessage id='$id' defaultMessage='$dataIndex' />, dataIndex: '$dataIndex', ellipsis: true,"

    if search_type == 'Select':
        column_str = column_str + "valueEnum: {0: {text: <FormattedMessage id='pages.common.true' />,status: 'True',},1: {text: <FormattedMessage id='pages.common.false' />,status: 'False',},}"
    elif search_type == '':
        column_str = column_str + ' search: false,'

    column_str = start_tag_str + column_str + end_tag_str
    return Template(column_str).substitute(lang_values)


def getTSInterface(ts_interface):
    ts_interface['type'] = 'string' if isinstance(ts_interface['value'], str) else 'undefined'
    interface_str = "$dataIndex : $type; "
    return Template(interface_str).substitute(ts_interface)


def genLang(sheet_name, table_olumns):
    sheet_names = '.'.join(sheet_name.lower().split('_'))
    # mock json
    mock_json = {}
    # columns ts interface
    ts_interface = ''
    # locales lang
    en_US_lang = {}
    zh_CN_lang = {}

    for column_item in table_olumns:
        colum_value = column_item['enValue'].title()
        colum_words = colum_value.split(' ')
        data_index = ''
        for index, word_item in enumerate(colum_words):
            if index == 0:
                word_item = word_item[:1].lower() + word_item[1:]
            data_index = data_index + word_item

        lang_key = 'pages.' + sheet_names + '.' + data_index
        mock_json[data_index] = colum_value
        en_US_lang[lang_key] = colum_value
        zh_CN_lang[lang_key] = colum_value if column_item['zhValue'] == '0' or column_item['zhValue'] == '' else \
            column_item['zhValue']

        # merge ts interface
        ts_interface_values = {'dataIndex': data_index, 'value': colum_value}
        ts_interface = ts_interface + getTSInterface(ts_interface_values)

    ts_interface = "{" + ts_interface + "}"
    #   mock data
    with open(write_path + "/" + sheet_names + '_' + "mock.json", "w", encoding="utf-8") as fen:
        fen.write(json.dumps(mock_json, indent=True, ensure_ascii=False))

    # columns ts interface
    with open(write_path + "/" + sheet_names + '_' + "interface.json", "w", encoding="utf-8") as fen:
        fen.write(json.dumps(ts_interface, indent=True, ensure_ascii=False))

    #   lang data
    with open(write_path + "/" + sheet_names + '_' + "EN.json", "w", encoding="utf-8") as fen:
        fen.write(json.dumps(en_US_lang, indent=True, ensure_ascii=False))
    with open(write_path + "/" + sheet_names + '_' + "ZH.json", "w", encoding="utf-8") as fen:
        fen.write(json.dumps(zh_CN_lang, indent=True, ensure_ascii=False))


def genCodes(sheet_name, form_type_data):
    sheet_names = '.'.join(sheet_name.lower().split('_'))
    # table columns
    columns_json = ''
    # form items
    form_items_json = ''

    for column_item in form_type_data:
        if 'enValue' in column_item:
            colum_value = column_item['enValue'].title()
            colum_words = colum_value.split(' ')
            data_index = ''
            for index, word_item in enumerate(colum_words):
                if index == 0:
                    word_item = word_item[:1].lower() + word_item[1:]
                data_index = data_index + word_item

            lang_key = 'pages.' + sheet_names + '.' + data_index

            # merge columns
            if 'searchType' in column_item:
                search_type = column_item['searchType']
                lang_values = {'id': lang_key, 'dataIndex': data_index}
                columns_json = columns_json + getTableTemplate(search_type, lang_values)

            # merge form items
            if 'formItemType' in column_item:
                form_item_type = column_item['formItemType']
                if form_item_type:
                    lang_labels = {'id': lang_key, 'type': form_item_type, 'name': data_index, }
                    form_items_json = form_items_json + getFormTemplate(lang_labels)

    columns_json = "[" + columns_json + "]"

    # columns
    with open(write_path + "/" + sheet_names + '_' + "columns.json", "w", encoding="utf-8") as fen:
        fen.write(json.dumps(columns_json, indent=True, ensure_ascii=False))

    # form items
    with open(write_path + "/" + sheet_names + '_' + "form_items.json", "w", encoding="utf-8") as fen:
        fen.write(json.dumps(form_items_json, indent=True, ensure_ascii=False))


def setLangData(sheet_name, sheet_content):
    try:
        en_data = dict(zip(sheet_content['KEY'], sheet_content['Value-EN']))
        zh_data = dict(zip(sheet_content['KEY'], sheet_content['Value-ZH']))
    except:
        print('Error:读取Search/Form KEY失败', sheet_name)
    else:
        columns_data = []
        for column_key in en_data:
            if column_key:
                columns_data.append({
                    'enValue': en_data[column_key],
                    'zhValue': zh_data[column_key],
                })
        # gen lang
        genLang(sheet_name, columns_data)


def setFormOrSearchData(sheet_name, sheet_content):
    try:
        en_data = dict(zip(sheet_content['KEY'], sheet_content['Value-EN']))
        search_type_data = dict(zip(sheet_content['KEY'], sheet_content['Search-Type']))
        form_type_data = dict(zip(sheet_content['KEY'], sheet_content['Form-Type']))
    except:
        print('Error:读取Search/Form KEY失败', sheet_name)
    else:
        columns_data = []
        for column_key in en_data:
            if column_key:
                columns_data.append({
                    'enValue': en_data[column_key],
                    'searchType': search_type_data[column_key],
                    'formItemType': form_type_data[column_key],
                })
        # gen codes
        genCodes(sheet_name, columns_data)


def readFiles():
    sheet_names = pd.read_excel(target_url, None).keys()
    for sheet_name in sheet_names:
        sheet_content = pd.read_excel(target_url, sheet_name=sheet_name, dtype=str, keep_default_na="")
        sheet_name = sheet_name[:1].lower() + sheet_name[1:]
        # 生成文件
        setLangData(sheet_name, sheet_content)
        setFormOrSearchData(sheet_name, sheet_content)


# run
readFiles()
