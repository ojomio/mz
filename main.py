from HTMLParser import HTMLParser
import argparse
from copy import copy
import re
from urllib import urlencode
import sys

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
            raise Exception('Hell! got %d', resp.text)
        return resp.json()['response']

    def upload_photo(self, filename, upload_url):
        files = {
            'file1': open(filename, 'rb')
        }
        resp = self.VKsession.post(upload_url, files=files)
        if resp.status_code != requests.codes.ok:
            raise Exception('Hell! got %d', resp.text)
        req_params = copy(resp.json())
        req_params['caption'] = filename
        req_params['description'] = filename
        return self.call_api('photos.save', req_params)


def main():
    parser = argparse.ArgumentParser(description='Upload files to VK group')
    parser.add_argument('--login', type=str)
    parser.add_argument('--password', type=str)
    parser.add_argument('--group_id', type=str, default='66887755')
    parser.add_argument('--album_id', type=str, default='188836852')
    parser.add_argument('--app_id', type=str, default='4203932')
    parser.add_argument('files', type=str, nargs='*')
    args = parser.parse_args()

    u = VKUploader(args.login, args.password, args.app_id)
    u.do_auth()

    payload = {
        'album_id': args.album_id,
        'group_id': args.group_id
    }
    upload_url = u.call_api('photos.getUploadServer', payload)['upload_url']
    print("Uploading photos to %s..." % upload_url)

    for (idx, filename) in enumerate(args.files, start=1):
        print('Uploading %d/%d file (%s)' % (idx, len(args.files), filename))
        print(u.upload_photo(filename, upload_url))


main()

