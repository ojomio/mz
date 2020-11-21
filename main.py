# coding=utf-8
import argparse
import io
import os
import re
import sys
from copy import copy
from html.parser import HTMLParser
from pprint import pprint
from threading import Event, Timer
from urllib.parse import urlencode

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
        self.method = "get"

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


def vk_api_ratelimited(func):
    def to_return(self, *args, **kwargs):
        def restriction_release(event_to_set):
            event_to_set.set()

        result = func(self, *args, **kwargs)
        self.api_requests_permitted.clear()
        Timer(0.3, restriction_release, args=[self.api_requests_permitted]).start()
        return result

    return to_return


class VKUploader(object):
    def __init__(self, app):
        self.app = app
        # No more than 3 reqs per second limit is enforced by this
        self.api_requests_permitted = Event()
        self.api_requests_permitted.set()

        self.do_auth()

    def do_auth(self):
        payload = {
            'client_id': self.app.app_id,
            'redirect_uri': "https://oauth.vk.com/blank.html",
            'v': '5.11',
            'scope': 4 + 262144 + 1048576,
            'display': 'wap',
            'response_type': 'token',
            'revoke': 1,
        }

        self.VKsession = requests.Session()
        resp_with_auth_form = self.VKsession.get(
            'https://oauth.vk.com/authorize', params=payload
        )
        parser = FormParser()
        parser.feed(resp_with_auth_form.text)

        req_params = copy(parser.params)
        req_params['email'] = self.app.login
        req_params['pass'] = self.app.password

        resp_with_grant_access_form = self.VKsession.__getattribute__(parser.method)(
            parser.url, params=req_params
        )
        parser = FormParser()
        parser.feed(resp_with_grant_access_form.text)

        resp_with_access_token = self.VKsession.__getattribute__(parser.method)(
            parser.url, params=parser.params
        )

        redir_location = resp_with_access_token.history[0].headers['Location']
        self.access_token = re.search(
            r'#.*access_token=(\w+)(&|$)', redir_location
        ).group(1)
        self.user_id = re.search(r'(?:#|&)user_id=(\w+)(&|$)', redir_location).group(1)

    @staticmethod
    def prepare_api_call_args(params):
        for (paramkey, paramvalue) in params.items():
            if isinstance(paramvalue, str):
                yield paramkey, paramvalue.encode('utf-8')
            else:
                yield paramkey, paramvalue

    @vk_api_ratelimited
    def call_api(self, api_method, params, http_verb='get'):
        # first ensure we waited enough to flood the server
        self.api_requests_permitted.wait()

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
        files = {'file1': (fileobj.filename, fileobj)}
        resp = self.VKsession.post(upload_url, files=files)
        if resp.status_code != requests.codes.ok:
            raise APIException('Hell! got %s' % resp.text)
        req_params = copy(resp.json())
        req_params['caption'] = description + '\n' + fileobj.filename
        req_params['description'] = description + '\n' + fileobj.filename
        return self.call_api('photos.save', req_params)

    def get_albums(self):
        req_params = {'owner_id': "-%s" % self.app.group_id}
        return {
            album_entry['title']: album_entry['aid']
            for album_entry in self.call_api('photos.getAlbums', req_params)
        }

    def create_album(self, title, descr=None):
        descr = descr or title
        req_params = {
            'group_id': "%s" % self.app.group_id,
            'title': title,
            'description': descr,
            'privacy': '0',
            'comment_privacy': '0',
        }
        return self.call_api('photos.createAlbum', req_params)['aid']

    def get_uploaded_goods_list(self, album):
        req_params = {'owner_id': "-%s" % self.app.group_id, 'album_id': album}
        return {
            self.extract_good_id(photo_entry): photo_entry['pid']
            for photo_entry in self.call_api('photos.get', req_params)
        }

    @staticmethod
    def extract_good_id(photo_entry):
        return os.path.splitext(photo_entry['text'].split('<br>')[-1])[0]


class Group(object):
    def __init__(self, app, album_id, goods_group_id, contained_goods):
        self.app = app
        self.id = goods_group_id
        self.uploaded_goods = self.app.u.get_uploaded_goods_list(album_id)
        payload = {'album_id': album_id, 'group_id': self.app.group_id}
        self.upload_url = self.app.u.call_api('photos.getUploadServer', payload)[
            'upload_url'
        ]
        self.contained_goods = contained_goods

    def sync_goods_in_group(self):
        print(("Uploading photos for album '%s' to %s..." % (self.id, self.upload_url)))
        self.delete_not_in_stock()
        self.add_not_in_album()

    def delete_not_in_stock(self):
        todel = set(self.uploaded_goods.keys()) - set(self.contained_goods.keys())
        print(('Removing %d goods not in stock...' % len(todel)))
        for (idx, good_id) in enumerate(todel, start=1):
            print(('Deleting %d/%d file (%s)' % (idx, len(todel), good_id)))
            self.app.u.call_api(
                'photos.delete',
                {
                    'owner_id': "-%s" % self.app.group_id,
                    'photo_id': self.uploaded_goods[good_id],
                },
            )

    def add_not_in_album(self):
        toadd = set(self.contained_goods.keys()) - set(self.uploaded_goods.keys())
        print(('Adding %d goods...' % len(toadd)))
        for (idx, good_id) in enumerate(toadd, start=1):
            sys.stdout.write("\r(%d/%d) " % (idx, len(toadd)))
            self.add_good(good_id, self.contained_goods[good_id])

    def add_good(self, good_id, description):
        try:
            fileobj = self.app.get_photo_from_site(good_id)
        except PhotoNotFound:
            try:
                fileobj = open('/'.join([self.app.missing_file_dir, good_id]), 'rb')
            except IOError:
                print(
                    ('%s was not found either on site or on disk . Skipping' % good_id)
                )
                return
        sys.stdout.write('Uploading %s/%s file' % (good_id, str(fileobj)))
        self.app.u.upload_photo(fileobj, self.upload_url, description)


class Application(object):
    def setup_arguments(self):
        parser = argparse.ArgumentParser(description='Upload files to VK group')
        parser.add_argument('--login', type=str)
        parser.add_argument('--password', type=str)
        parser.add_argument('--group_id', type=str, default='68712727')
        parser.add_argument('--app_id', type=str, default='4203932')
        parser.add_argument('--missing-file-dir', type=str, default='.')
        parser.add_argument('stock_list', type=str)
        parser.parse_args(namespace=self)

    def __init__(self):
        self.u = None
        self.stock_list = None
        self.password = None
        self.login = None

    def create_missing_albums(self, goods_in_stock, existing_albums):
        albums_to_create = set(goods_in_stock.keys()) - set(existing_albums.keys())
        print(('Creating missing photo albums(%d)...' % len(albums_to_create)))
        for album_name in albums_to_create:
            print('Creating %s...' % album_name)
            existing_albums[album_name] = self.u.create_album(album_name)

    @staticmethod
    def get_goods_in_stock(stock_list_file):
        def entries_():
            for row_number in range(sh.nrows):
                cell_value = str(sh.cell_value(row_number, 0))
                if not is_good_id(cell_value):
                    continue
                yield (
                    {
                        'id': str(int(float(cell_value.strip()))),
                        'group_id': sh.cell_value(row_number, 4),
                        'description': '\n'.join(
                            sh.cell_value(row_number, i) for i in (1, 2, 8)
                        )
                        # Наименование, класс, характеристики
                    }
                )

        goods_by_group = {}

        with xlrd.open_workbook(stock_list_file) as wb:
            sh = wb.sheet_by_index(0)
            for entry in entries_():
                group_dict = goods_by_group.get(entry['group_id'], {})
                group_dict[entry['id']] = entry['description']
                goods_by_group[entry['group_id']] = group_dict
        del goods_by_group['NULL']  # Не показывать эту группу
        return goods_by_group

    @staticmethod
    def get_photo_from_site(filename):
        filename = os.path.basename(filename)
        sys.stdout.write('Getting photo for goods id %s... ' % filename)
        resp = requests.get('http://texrepublic.ru/pic/site/%s.gif' % filename)

        if resp.status_code != requests.codes.ok:
            raise PhotoNotFound('Hell! got %s', resp.text)

        ret = io.StringIO(resp.content)
        ret.filename = '%s.gif' % filename
        return ret


def main():
    app = Application()
    app.setup_arguments()

    u = VKUploader(app)
    app.u = u
    existing_albums = u.get_albums()
    goods_in_stock = app.get_goods_in_stock(app.stock_list)
    app.create_missing_albums(goods_in_stock, existing_albums)
    for goods_group_id in goods_in_stock.keys():
        g = Group(
            album_id=existing_albums[goods_group_id],
            contained_goods=goods_in_stock[goods_group_id],
            goods_group_id=goods_group_id,
            app=app,
        )
        g.sync_goods_in_group()


def is_good_id(str_):
    return bool(re.match(r'^[\d\s.]+$', str_))


main()
