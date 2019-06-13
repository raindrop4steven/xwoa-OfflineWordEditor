#!/usr/bin/env python
# encoding: utf-8

import pdb
import os
import logging
import json
import subprocess
import requests
from concurrent.futures import ThreadPoolExecutor as Pool
from flask import Flask, jsonify, request

# Constants defined here
DOWNLOAD_ATT_FOLDER = r'C:\AttFolder'


app = Flask(__name__)
pool = Pool()
logging.basicConfig(level=logging.INFO)

future_dict = {}
file_status_dict = {}

@app.route('/')
def index():
    return jsonify({
        "status": "running..."
    })

@app.route('/status/<string:att_id>')
def status(att_id):
    # Check attachment editing status
    att_status = file_status_dict.pop(att_id, None)
    if att_status is None:
        return jsonify({
            'code': 201,
            'message': '处理中'
        })
    elif att_status.get('status') == 'uploaded':
        return jsonify({
            'code': 200,
            'message': '上传成功'
        })
    elif att_status.get('status') == 'error':
        return jsonify({
            'code': 400,
            'message': att_status.get('message', '')
        })
    else:
        return jsonify({
            'code': 202,
            'message': '处理中'
        })

@app.route('/openDoc', methods=['POST'])
def openDoc():
    # Parse arguments
    origin = request.form.get('origin')
    session = request.form.get('session')
    att_id = request.form.get('att_id')
    extension = request.form.get('extension')
    row = request.form.get('row')
    col = request.form.get('col')

    # check if file already opened.
    for future_item in future_dict.values():
        if future_item['att_id'] == att_id:
            # already opened files.
            return jsonify({
                'code': 400,
                'message': '文档已被打开'
            })
        else:
            continue
    # download file
    file_path = download_file(origin, session, att_id, extension)

    f = pool.submit(subprocess.call, r'start /wait %s' % file_path, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, stdin=subprocess.PIPE)
    f.add_done_callback(callback)
    future_dict[f] = {
        'att_id': att_id,
        'file_path': file_path,
        'session': session,
        'origin': origin,
        'row': row,
        'col': col
    }
    # pool.shutdown(wait=True)

    return jsonify({
        'code': 200
    }) 

def after_request(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'PUT,GET,POST,DELETE'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type,Authorization'
    return response

def download_file(origin, session, att_id, extension):
    # TODO: local storage
    local_filename = f'{att_id}.{extension}'
    file_path = os.path.join(DOWNLOAD_ATT_FOLDER, local_filename)

    url = f'{origin}/Apps/Workflow/Worksheet/DownloadAtt/{att_id}?__session__={session}'
    # NOTE the stream=True parameter below
    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        with open(file_path, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192): 
                if chunk: # filter out keep-alive new chunks
                    f.write(chunk)
                    
    return file_path

def callback(future):
    if future.exception() is not None:
        logging.error('got exception: %s' % future.exception())
        pass
    else:
        process_dict = future_dict.pop(future, None)
        if process_dict is None:
            logging.error('process dict not found future')
        else:
            logging.debug('process returned %d' % future.result())
            upload_modified_file(process_dict)

def upload_modified_file(p_dict):
    # save cookie into local disk file.
    att_id = p_dict['att_id']
    file_path = p_dict['file_path']
    session = p_dict['session']
    origin = p_dict['origin']
    row = p_dict['row']
    col = p_dict['col']
    
    files = {'file': open(file_path, 'rb')}
    
    r = requests.post(f'{origin}/Apps/DEP/Common/Upload?__session__={session}', data={'AttID': att_id, 'row': row, 'col': col}, files=files)
    json_data = json.loads(r.content)
    if json_data['Succeed'] is True:
        logging.info('上传成功')
        file_status_dict[att_id] = {'status': 'uploaded', 'message': ''}
    else:
        err_msg = f'上传失败: {json_data["Message"]}'
        logging.error(err_msg)
        file_status_dict[att_id] = {'status': 'error', 'message': err_msg}


if __name__ == '__main__':
    # Check if att folder exists
    if not (os.path.isdir(DOWNLOAD_ATT_FOLDER)):
        # create att folder if not exists
        os.mkdir(DOWNLOAD_ATT_FOLDER)
    # disable logger
    # log = logging.getLogger('werkzeug')
    # log.disabled = True
    # app.logger.disabled = True
    app.after_request(after_request)
    app.run(debug=True)
