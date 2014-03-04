from HTMLParser import HTMLParser
import StringIO
import argparse
from copy import copy
import os
import re
from urllib import urlencode

import xlrd
from xlrd.xlsx import cell_name_to_rowx_colx


__author__ = 'crystal'

import requests


class FormParser(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self)
        self.url = None
        self.params = {}
        self.in_form = False
        self.form_parsed = False
        self.method = "GET"

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == "form":
            if self.form_parsed:
                raise RuntimeError("Second form on page")
            if self.in_form:
                raise RuntimeError("Already in form")
            self.in_form = True
        if not self.in_form:
            return
        attrs = dict((name.lower(), value) for name, value in attrs)
        if tag == "form":
            self.url = attrs["action"]
            if "method" in attrs:
                self.method = attrs["method"]
        elif tag == "input" and "type" in attrs and "name" in attrs:
            if attrs["type"] in ["hidden", "text", "password"]:
                self.params[attrs["name"]] = attrs["value"] if "value" in attrs else ""

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag == "form":
            if not self.in_form:
                raise RuntimeError("Unexpected end of <form>")
            self.in_form = False
            self.form_parsed = True


class APIException(Exception):
    pass


class PhotoNotFound(Exception):
    pass


class VKUploader(object):
    def __init__(self, login, password, app_id):
        self.login = login
        self.password = password
        self.app_id = app_id

    def do_auth(self):
        payload = {
            'client_id': self.app_id,
            'redirect_uri': "https://oauth.vk.com/blank.html",
            'v': '5.11',
            'scope': "1310724",
            'display': 'wap',
            'response_type': 'token',
            'revoke': 1,
        }

        self.VKsession = requests.Session()
        resp_with_auth_form = self.VKsession.get('https://oauth.vk.com/authorize', params=payload)
        parser = FormParser()
        parser.feed(resp_with_auth_form.text)

        req_params = copy(parser.params)
        req_params['email'] = self.login
        req_params['pass'] = self.password

        resp_with_grant_access_form = \
            self.VKsession.__getattribute__(parser.method)(parser.url, params=req_params)
        parser = FormParser()
        parser.feed(resp_with_grant_access_form.text)

        resp_with_access_token = \
            self.VKsession.__getattribute__(parser.method)(parser.url, params=parser.params)

        redir_location = resp_with_access_token.history[0].headers['Location']
        self.access_token = re.search(r'#.*access_token=(\w+)(&|$)', redir_location).group(1)
        self.user_id = re.search(r'(?:#|&)user_id=(\w+)(&|$)', redir_location).group(1)


    def call_api(self, method, params, http_verb='get'):
        if self.access_token:
            params["access_token"] = self.access_token
        url = 'https://api.vk.com/method/%s?%s' % (method, urlencode(params))

        resp = self.VKsession.__getattribute__(http_verb)(url)
        if resp.status_code != requests.codes.ok:
            raise APIException('Hell! got %s', resp.text)
        return resp.json()['response']

    def upload_photo(self, fileobj, upload_url):
        files = {
            'file1': (fileobj.filename, fileobj)
        }
        resp = self.VKsession.post(upload_url, files=files)
        if resp.status_code != requests.codes.ok:
            raise Exception('Hell! got %s', resp.text)
        req_params = copy(resp.json())
        req_params['caption'] = fileobj.filename
        req_params['description'] = fileobj.filename
        return self.call_api('photos.save', req_params)

    def get_uploaded_goods_list(self):
        req_params = {
            'owner_id': "-%s" % args.group_id,
            'album_id': args.album_id,
        }
        return {
            os.path.splitext(photo_entry['text'])[0]: photo_entry['pid']
            for photo_entry
            in self.call_api('photos.get', req_params)
        }


def get_photo_from_site(filename):
    filename = os.path.basename(filename)
    print('Getting photo for goods id %s ...' % filename)
    resp = requests.get('http://texrepublic.ru/pic/site/%s.gif' % filename)

    if resp.status_code != requests.codes.ok:
        raise PhotoNotFound('Hell! got %s', resp.text)

    ret = StringIO.StringIO(resp.content)
    ret.filename = '%s.gif' % filename
    return ret


def add_good(good_id, idx, total, uploader, upload_url):
    try:
        fileobj = get_photo_from_site(good_id)
    except PhotoNotFound:
        try:
            fileobj = open('/'.join([args.missing_file_dir, good_id]), 'rb')
        except IOError:
            print('%s was not found either on site or on disk . Skipping' % good_id)
            return

    print('Uploading %d/%d file (%s/%s)' % (idx, total, good_id, str(fileobj)))
    uploader.upload_photo(fileobj, upload_url)


def is_good_id(str_):
    return bool(re.match(r'^(\d|\s)+$', str_))


def get_goods_in_stock(stock_list_file):
    with xlrd.open_workbook(stock_list_file) as wb:
        sh = wb.sheet_by_index(0)
        for cell_value in sh.col_values(1):
            if not is_good_id(cell_value):
                continue

            yield cell_value.strip()


def main():
    parser = argparse.ArgumentParser(description='Upload files to VK group')
    parser.add_argument('--login', type=str)
    parser.add_argument('--password', type=str)
    parser.add_argument('--group_id', type=str, default='66887755')
    parser.add_argument('--album_id', type=str, default='188836852')
    parser.add_argument('--app_id', type=str, default='4203932')
    parser.add_argument('--missing-file-dir', type=str, default='.')

    parser.add_argument('stock_list', type=str)
    global args
    args = parser.parse_args()

    u = VKUploader(args.login, args.password, args.app_id)
    u.do_auth()

    payload = {
        'album_id': args.album_id,
        'group_id': args.group_id
    }
    upload_url = u.call_api('photos.getUploadServer', payload)['upload_url']
    print("Uploading photos to %s..." % upload_url)

    uploaded_goods = u.get_uploaded_goods_list()
    goods_in_stock = set(get_goods_in_stock(args.stock_list))
    todel = set(uploaded_goods.keys()) - goods_in_stock

    print ('Removing %d goods not in stock...' % len(todel))
    for (idx, good_id) in enumerate(todel, start=1):
        print('Deleting %d/%d file (%s)' % (idx, len(todel), good_id))
        try:
            u.call_api('photos.delete', {'owner_id': u.user_id, 'photo_id': uploaded_goods[good_id]})
        except APIException as e:
            print('Error deleting: %s' % str(e))

    toadd = goods_in_stock - set(uploaded_goods.keys())
    print ('Adding %d goods...' % len(toadd))
    for (idx, good_id) in enumerate(toadd, start=1):
        add_good(good_id, idx, len(toadd), u, upload_url)


main()

