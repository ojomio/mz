# coding=utf-8
from HTMLParser import HTMLParser
import StringIO
import argparse
from copy import copy
import os
import re
from urllib import urlencode

import xlrd


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
        self.do_auth()

    def do_auth(self):
        payload = {
            'client_id': self.app_id,
            'redirect_uri': "https://oauth.vk.com/blank.html",
            'v': '5.11',
            'scope': 4 + 262144 + 1048576,
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


    @staticmethod
    def prepare_api_call_args(params):
        for (paramkey, paramvalue) in params.iteritems():
            if isinstance(paramvalue, unicode):
                yield paramkey, paramvalue.encode('utf-8')
            else:
                yield paramkey, paramvalue

    def call_api(self, api_method, params, http_verb='get'):
        if self.access_token:
            params["access_token"] = self.access_token
        params = dict(self.prepare_api_call_args(params))
        url = 'https://api.vk.com/method/%s?%s' % (api_method, urlencode(params))

        resp = self.VKsession.__getattribute__(http_verb)(url)

        if resp.status_code != requests.codes.ok:
            raise APIException('Hell! got %s', resp.text)
        json = resp.json()
        if 'error' in json:
            raise APIException('Hell! got %s' % json)

        return json['response']

    def upload_photo(self, fileobj, upload_url, description):
        files = {
            'file1': (fileobj.filename, fileobj)
        }
        resp = self.VKsession.post(upload_url, files=files)
        if resp.status_code != requests.codes.ok:
            raise APIException('Hell! got %s' % resp.text)
        req_params = copy(resp.json())
        req_params['caption'] = description + '\n' + fileobj.filename
        req_params['description'] = description + '\n' + fileobj.filename
        return self.call_api('photos.save', req_params)

    def get_albums(self):
        req_params = {
            'owner_id': "-%s" % args.group_id,
        }
        return {
            album_entry['title']: album_entry['aid']
            for album_entry
            in self.call_api('photos.getAlbums', req_params)
        }

    def create_album(self, title, descr=None):
        descr = descr or title
        req_params = {
            'group_id': "%s" % args.group_id,
            'title': title,
            'description': descr,
            'privacy': '0',
            'comment_privacy': '0',

        }
        return self.call_api('photos.createAlbum', req_params)['aid']

    def get_uploaded_goods_list(self, album):
        req_params = {
            'owner_id': "-%s" % args.group_id,
            'album_id': album,
        }
        return {
            self.extract_good_id(photo_entry): photo_entry['pid']
            for photo_entry
            in self.call_api('photos.get', req_params)
        }

    @staticmethod
    def extract_good_id(photo_entry):
        return os.path.splitext(
            photo_entry['text'].split('<br>')[-1]
        )[0]



def main():
    setup_arguments()

    u = VKUploader(args.login, args.password, args.app_id)

    existing_albums = u.get_albums()
    goods_in_stock = get_goods_in_stock(args.stock_list)
    create_missing_albums(u, existing_albums, goods_in_stock)
    for goods_group_id in goods_in_stock.iterkeys():
        sync_goods_in_group(existing_albums, goods_group_id, goods_in_stock, u)


def setup_arguments():
    parser = argparse.ArgumentParser(description='Upload files to VK group')
    parser.add_argument('--login', type=str)
    parser.add_argument('--password', type=str)
    parser.add_argument('--group_id', type=str, default='66887755')
    parser.add_argument('--app_id', type=str, default='4203932')
    parser.add_argument('--missing-file-dir', type=str, default='.')
    parser.add_argument('stock_list', type=str)
    global args
    args = parser.parse_args()


def get_goods_in_stock(stock_list_file):
    def entries_():
        for row_number in range(sh.nrows):
            cell_value = unicode(sh.cell_value(row_number, 1))
            if not is_good_id(cell_value):
                continue

            yield (
                {
                    'id': str(int(float(cell_value.strip()))),
                    'group_id': sh.cell_value(row_number, 5),
                    'description': '\n'.join(sh.cell_value(row_number, i) for i in (2, 3, 9))
                    # Наименование, класс, характеристики
                }
            )

    goods = {}

    with xlrd.open_workbook(stock_list_file) as wb:
        sh = wb.sheet_by_index(0)
        for entry in entries_():
            group_dict = goods.get(entry['group_id'], {})
            group_dict[entry['id']] = entry['description']
            goods[entry['group_id']] = group_dict
    del goods['NULL']  # Не показывать эту группу
    return goods


def create_missing_albums(u, goods_in_stock, existing_albums):
    albums_to_create = set(goods_in_stock.keys()) - set(existing_albums.keys())
    print('Creating missing photo albums(%d)...' % len(albums_to_create))
    for album_name in albums_to_create:
        print 'Creating %s...' % album_name
        existing_albums[album_name] = u.create_album(album_name)


def sync_goods_in_group(existing_albums, goods_group_id, goods_in_stock, u):
    payload = {
        'album_id': existing_albums[goods_group_id],
        'group_id': args.group_id
    }
    upload_url = u.call_api('photos.getUploadServer', payload)['upload_url']
    print("Uploading photos for album '%s' to %s..." % (goods_group_id, upload_url))

    uploaded_goods = u.get_uploaded_goods_list(existing_albums[goods_group_id])
    delete_not_in_stock(goods_in_stock[goods_group_id], u, uploaded_goods)
    add_not_in_album(goods_in_stock[goods_group_id], u, upload_url, uploaded_goods)


def delete_not_in_stock(goods_in_group, u, uploaded_goods):
    todel = set(uploaded_goods.keys()) - set(goods_in_group.keys())
    print('Removing %d goods not in stock...' % len(todel))
    for (idx, good_id) in enumerate(todel, start=1):
        print('Deleting %d/%d file (%s)' % (idx, len(todel), good_id))
        u.call_api('photos.delete', {'owner_id': "-%s" % args.group_id, 'photo_id': uploaded_goods[good_id]})


def add_not_in_album(goods_in_group, u, upload_url, uploaded_goods):
    toadd = set(goods_in_group.keys()) - set(uploaded_goods.keys())
    print('Adding %d goods...' % len(toadd))
    for (idx, good_id) in enumerate(toadd, start=1):
        add_good(good_id, idx, len(toadd), u, upload_url, goods_in_group[good_id])


def add_good(good_id, idx, total, uploader, upload_url, description):
    try:
        fileobj = get_photo_from_site(good_id)
    except PhotoNotFound:
        try:
            fileobj = open('/'.join([args.missing_file_dir, good_id]), 'rb')
        except IOError:
            print('%s was not found either on site or on disk . Skipping' % good_id)
            return

    print('Uploading %d/%d file (%s/%s)' % (idx, total, good_id, str(fileobj)))
    uploader.upload_photo(fileobj, upload_url, description)


def get_photo_from_site(filename):
    filename = os.path.basename(filename)
    print('Getting photo for goods id %s ...' % filename)
    resp = requests.get('http://texrepublic.ru/pic/site/%s.gif' % filename)

    if resp.status_code != requests.codes.ok:
        raise PhotoNotFound('Hell! got %s', resp.text)

    ret = StringIO.StringIO(resp.content)
    ret.filename = '%s.gif' % filename
    return ret


def is_good_id(str_):
    return bool(re.match(r'^[\d\s.]+$', str_))


main()
