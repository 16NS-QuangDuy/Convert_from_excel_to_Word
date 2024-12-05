# encoding: utf-8
import os
import requests
from bs4 import BeautifulSoup
from os.path import exists
import json


class CurlML:
    """CurlML"""
    auto_worker_name = "CurlML"

    @staticmethod
    def upload_file_confluence2(filename, destination='https://insight.fsoft.com.vn/confluence/display/XHSCHC32A4A0/EBTresosPluginsBuilder'):
        output_dir = os.path.dirname(filename)
        FileList = CurlML.split_file(filename, output_dir)
        for filename_i in FileList:
            CurlML.upload_file_confluence(filename_i, destination)

    @staticmethod
    def upload_file_confluence(file,
                               destination='https://insight.fsoft.com.vn/confluence/display/XHSCHC32A4A0/EBTresosPluginsBuilder'):
        """
            Use module to filter data
        """
        if (exists(file)):
            username = 'haindm'
            password = 'UMC4X@12345678'
            filePath = file.replace('\\', '/')
            fileName = filePath.split('/')[-1]
            session = requests.Session()
            login_data = {'os_username': username, 'os_password': password, 'login': 'Log in', 'os_destination': ''}
            print('Logging into confluence...\n')
            response = session.post('https://insight.fsoft.com.vn/confluence/dologin.action', login_data)
            print('Sign-in successfully.\n')
            pageId = BeautifulSoup(session.get(destination).content, 'html.parser').find('meta',
                                                                                         attrs={"name": "ajs-page-id",
                                                                                                "content": True})[
                'content']
            pageLink = 'https://insight.fsoft.com.vn/confluence/rest/api/content/' + str(pageId) + '/child/attachment'
            print('PageId=' + str(pageId) + '\n')
            headers = {'X-Atlassian-Token': 'no-check', }
            files = {'file': open(filePath, 'rb'), }
            print('Upload file...\n')
            response = session.post(pageLink, headers=headers, files=files, auth=('admin', 'admin'))
            if (400 == response.status_code):
                response_dict = json.loads(response.text)
                if (response_dict['message'].startswith(
                        'Cannot add a new attachment with same file name as an existing attachment')):
                    print('File already exists!\n')
                    print('Update new version for attachment...\n')
                    file_infor_tag = BeautifulSoup(session.get(destination).content, 'html.parser').find('tr', attrs={
                        "data-attachment-filename": fileName, "class": True, "data-attachment-id": True})
                    if (None != file_infor_tag):
                        fileId = file_infor_tag["data-attachment-id"]
                        pageLink += '/' + str(fileId) + '/data'
                        files = {'file': open(filePath, 'rb'), }
                        response = session.post(pageLink, headers=headers, files=files, auth=('admin', 'admin'))
                        response_dict = json.loads(response.text)
                        print('Upload successfully!\n')
                        print('File name: ' + response_dict['title'] + ', version: ' + str(
                            response_dict['version']['number']) + ', by ' + response_dict['version']['by'][
                                  'displayName'] + ', ' + response_dict['version']['when'])
                else:
                    print(response_dict['message'])
            else:
                response_dict = json.loads(response.text)
                print('Upload successfully!\n')
                print('File name: ' + response_dict['results'][0]['title'] + ', version: ' + str(
                    response_dict['results'][0]['version']['number']) + ', by ' +
                      response_dict['results'][0]['version']['by']['displayName'] + ', ' +
                      response_dict['results'][0]['version']['when'])
        else:
            print('File is not exists\n')

    @staticmethod
    def split_file(filename, output_dir):
        chunk_size = 25000000
        FileList = []
        if os.path.getsize(filename) > chunk_size:
            with open(filename, 'rb') as f:
                chunk = f.read(chunk_size)
                count = 0
                while chunk:
                    count += 1
                    suffix = '0' * (3 - len(str(count))) + str(count)
                    new_file_name = "%s.%s" % (filename, suffix)
                    new_file_name = os.path.join(output_dir, new_file_name)
                    with open(new_file_name, 'wb') as fout:
                        fout.write(chunk)
                        fout.close()
                    FileList.append(new_file_name)
                    chunk = f.read(chunk_size)
        else:
            FileList.append(filename)
        return FileList

    @staticmethod
    def download_file_confluence(link, destination, username='datnq9', password='Muathu2022@12345'):
        """
            download file from confluence to output folder
            link: file link
            destimation: output folder
        """
        # username = 'datnq9'
        # password = 'Muathu2022@12345'
        filename = link.replace('\\', '/').split('/')[-1].replace('?api=v2', '').replace('%20', ' ')
        filePath = destination.replace('\\', '/').rstrip('/') + '/' + filename
        session = requests.Session()
        login_data = {'os_username': username, 'os_password': password, 'login':'Log in', 'os_destination':''}
        print('Logging into confluence...\n')
        response = session.post('https://insight.fsoft.com.vn/confluence/dologin.action', login_data)
        if(None != BeautifulSoup(response.content , 'html.parser').find('div', attrs={"id" : "login-messages", "class" : "messages error"})):
            print('Sign-in failed!\n')
        else:
            print('Sign-in successfully.\n')
            print('Download file...\n')
            response = session.get(link.replace('\\', '/'))
            open(filePath, 'wb').write(response.content)
            print('Done.\n')

