# using flask_restful
from flask import Flask, jsonify, request, send_file, after_this_request
from flask_restful import Resource, Api
from flask_cors import CORS

# pyopenxl
from openpyxl import Workbook

# utilities
import os
import random
import string
import argparse


MINE_TYPE_FOR_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def get_random_string(length):
    # choose from all lowercase letter
    letters = string.ascii_lowercase
    result_str = ''.join(random.choice(letters) for i in range(length))
    return result_str


def get_data_and_delete_file(file_name):
    file = open(file_name, "rb")
    data = file.read()
    file.close()
    os.remove(file_name)  # clean tmp file
    return data


# creating the flask app
app = Flask(__name__)
# Enable CORS
CORS(app)
# creating an API object
api = Api(app)

# making a class for a particular resource
# the get, post methods correspond to get and post requests
# they are automatically mapped by flask_restful.
# other methods include put, delete, etc.


class Hello(Resource):
    # corresponds to the GET request.
    # this function is called whenever there
    # is a GET request for this resource
    def get(self):
        return jsonify({'message': 'hello world!!'})

    # Corresponds to POST request
    def post(self):
        data = request.get_json()	 # status code
        return jsonify({'data': data}), 201


# another resource to calculate the square of a number
class Square(Resource):
    def get(self, num):
        return jsonify({'square': num**2})


class Excel(Resource):
    def post(self):
        post_data = request.get_json()

        def generate():
            xl_file_name = get_random_string(10) + ".xlsx"
            wb = Workbook()
            ws = wb.active
            ws['A1'] = post_data["luong"]
            wb.save(xl_file_name)

            data = get_data_and_delete_file(xl_file_name)
            yield data

        return app.response_class(generate(), mimetype=MINE_TYPE_FOR_XLSX)


# adding the defined resources along with their corresponding urls
api.add_resource(Hello, '/')
api.add_resource(Square, '/square/<int:num>')
api.add_resource(Excel, '/excel')


def make_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--ip", default="0.0.0.0",
                        help="Ip of running host")
    parser.add_argument("-p", "--port", default=4444, help="Port of host")
    return parser


args = make_parser().parse_args()

# driver function
if __name__ == '__main__':
    app.run(debug=True, host=args.ip, port=args.port)
