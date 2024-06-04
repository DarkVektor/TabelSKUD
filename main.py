import json
import os
import openpyxl

import requests
from openpyxl.workbook import Workbook


def GetToken(_config):
    _token = ''
    try:
        _url = f'http://{_config["serverIP"]}/api/system/auth'
        _payload = {
            "login": _config['login'],
            "password": _config['password']
        }
        _headers = {
            "Content-type": "application/json; charset=UTF-8",
            #"Authorization": "Bearer rJByeM5PJanwBErcGAtIoBG5CitfQFNj"
        }
        _token = requests.request("post", _url, json=_payload, headers=_headers, timeout=0.2).json()['token']
    except Exception as e:
        print(f'Ошибка при авторизации: {e}')
    return _token

def GetAccessReport(_config, _token):

    def ResponceToString(_responce, _config):
        _answer_string = ''
        for user in responce['rows']:
            flag = True
            for word in _config['wrong_words']:
                if word in user['fio']:
                    flag = False
                    break
            if flag:
                _answer_string += f'{user["fio"]}|{user["position_name"]}|{user["division_name"]}|{user["zone_exit"]}|{user["zone_enter"]}|{user["time_label"]}\n'
        return _answer_string

    answer = ''
    try:
        _url = f'http://{_config["serverIP"]}/api/accessReports/events'
        payload = {
            'token': _token,
            'page': 1,
            'sord': 'asc',
            'dateBegin': _config["dateBegin"],
            'dateEnd': _config["dateEnd"],
            'searchString': 'КПП производства'
        }
        headers = {
            "Content-type": "application/json; charset=UTF-8",
            # "Authorization": "Bearer rJByeM5PJanwBErcGAtIoBG5CitfQFNj"
        }
        responce = requests.request('get', _url, params=payload, headers=headers).json()
        count_pages = responce['total']
        print(f'Страница {payload["page"]} из {count_pages}')
        answer += ResponceToString(responce, _config)
        if count_pages != 0:
            for i in range(count_pages - 1):
                payload['page'] += 1
                responce = requests.request('get', _url, params=payload, headers=headers).json()
                print(f'Страница {payload["page"]} из {count_pages}')
                answer += ResponceToString(responce, _config)
    except Exception as e:
        print(f'Ошибка при запросе данных о проходе')
    return answer

def StringToArray(_accessReport):
    _array = list()
    _userList = list()
    _user = ''
    _accessArray = _accessReport.split('\n')
    for access in _accessArray:
        row = access.split('|')
        if _user == row[0]:
            _userList.append(row)
        else:
            if len(_userList) != 0:
                _userList = sorted(_userList, key=lambda x: x[5], reverse=False)
                _array.append(_userList)
            _userList = list()
            _userList.append(row)
            _user = row[0]
    return _array

def SaveReports(_arrayAccessReports):
    try:
        for _user in _arrayAccessReports:
            wb = Workbook()
            ws = wb.active
            _name = _user[0][0]
            print(_name)
            for _enter in _user:
                row = (_enter)
                ws.append(row)
            wb.save("Информация о проходах/" + _name + ".xlsx")
    except Exception as e:
        print('При заполнении таблицы о проходах возникла ошибка')

if __name__ == '__main__':
    #Загрузка параметров работы программы
    try:
        with open("config.json", 'r', encoding='ANSI') as file:
            config = json.load(file)
    except:
        with open("config.json", 'w', encoding='ANSI') as file:
            config = dict()
            config['serverIP'] = '192.168.9.52'
            config['login'] = 'admin'
            config['password'] = 'adminfarm123'
            config['dateBegin'] = '2024.06.01'
            config['dateEnd'] = '2024.12.31'
            config['wrong_words'] = ["Полный", "1", "кпп"]
            json.dump(config, file, indent=4)
    tk = GetToken(config)
    #При отсутствии токена, завершается программа
    if tk == '':
        print('Не получен токен авторизации. Программа завершена с ошибкой')
        os.kill(os.getpid(),9)
    accessReport = GetAccessReport(config, tk)
    #print(accessReport)
    arrayAccessReport = StringToArray(accessReport)
    SaveReports(arrayAccessReport)