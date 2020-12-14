# TME API - example usage with Python
# More info and create application at: https://developers.tme.eu
# Пример хеширования на Python3: https://stackoverflow.com/questions/37763235/unicode-objects-must-be-encoded-before-hashing-error
# hmac - Keyed-Hashing for Message Authentication
# Интересная статья https://learn.javascript.ru/xhr-forms
import collections, urllib, base64, hmac, hashlib, json # urllib2
import urllib.request

def product_import_tme(token, app_secret, action, params):
    # /product/product_import_tme/
    token = token.encode() # TOKEN
    app_secret = app_secret.encode() #bytes('0b748f6e5d340d693703', encoding='utf8') #App secret
    response = api_call( action, params, token, app_secret, True)
    response = json.loads(response)
    
    return response

def api_call(action, params, token, app_secret, show_header=False):
    api_url = 'https://api.tme.eu/' + action + '.json'# URI + request method
    params['Token'] = token
    # Сортировка: элементы запроса должны быть отсортированы в алфавитном порядке до шифрования
    params = collections.OrderedDict(sorted(params.items()))
    encoded_params = urllib.parse.urlencode(params, quote_via=urllib.parse.quote)# было ",'' )" это для артикулов с пробелом "quote_via=urllib.parse.quote"
    #print(encoded_params)
    # СИГНАТУРА ЗАПРОСА:
    # Базовая последовательность: Метод HTTP-запроса (POST, GET, HEAD) должны быть написанны заглавными буквами 
    # МЕТОД + & (код ASCII:38) + Базовая последовательность URI + & + Стандартизированные параметры запроса после шифрования urlencoded HTTP
    signature_base = 'POST' + '&' + urllib.parse.quote(api_url, '') + '&' + urllib.parse.quote(encoded_params, '')#,safe=''
    #Сгенерироваnm подпись HMAC-SHA1 для закрытого ключа средствами функции Base64
    api_signature = base64.encodebytes(hmac.new(app_secret, signature_base.encode(), hashlib.sha1).digest()).rstrip() # проблемы с кодировкой в оригинальной строке
    #api_signature = base64.b64encode(hmac.new(app_secret, signature_base.encode(), hashlib.sha1).digest()) # подпись HMAC-SHA1, сгенерированная в двоичной форме
    params['ApiSignature'] = api_signature
    params = collections.OrderedDict(sorted(params.items())) # еще раз отсортировать на всякий
    #opts = {'http': {'method' : 'POST','header' : 'Content-type: application/x-www-form-urlencoded','content' : urllib.parse.urlencode(params)}}
    # браузер кодирует данные соответствующим способом перед отправкой на сервер, для метода POST доступны три кодировки: application/x-www-form-urlencoded * multipart/form-data * text-plain
    # при POST обязателен заголовок Content-Type, содержащий кодировку. Это указание для сервера – как обрабатывать (раскодировать) пришедший запрос.
    http_header = {"Content-type": "application/x-www-form-urlencoded",} 
    # create your HTTP request
    req = urllib.request.Request(api_url, urllib.parse.urlencode(params).encode('utf-8'), http_header) # OK 
    # submit your request
    res = urllib.request.urlopen(req)
    html = res.read()
    #inform = res.info()
    return html

if __name__ == "__main__":
    token = 'ac434c181917ed4e51c49a2027bfd040e9f2da0054be7'
    app_secret = '0b748f6e5d340d693703'
    action = 'Products/GetSymbols' # request method
    params = {
            'CategoryId' : '100544',
            'Country' : 'RU',
            'Language' : 'RU',
        }
    print(product_import_tme(token, app_secret,action,params))