import uno
from wordconverter import *
from os.path import isfile
from sys import argv, exit

from flask import Flask, jsonify, make_response, request, send_file, flash, redirect
from werkzeug.utils import secure_filename
import logging
from gevent.wsgi import WSGIServer

app = Flask(__name__)

converter = DocumentConverter()

@app.route('/health', methods=['GET'])
def health():
	return jsonify({"message": "I AM ALIVE!!"})

@app.route("/api/convert", methods=['POST'])
def main():

	#check content length of request
	content_length = int(request.headers.get('Content-length'))
	if content_length == 0:
		error = "Request data has zero length."
		app.logger.error(error)
		return make_response(jsonify({'error': error}), 411)

	# check if the post request has the file part
	if 'file' not in request.files:
		error = "Please send a file."
		app.logger.error(error)
		return make_response(jsonify({'error': error}), 412)

	input_file = request.files['file']

	# if user does not select file, browser also submit a empty part without filename
	if input_file.filename == '':
		error = "Filename is empty."
		app.logger.error(error)
		return make_response(jsonify({'error': error}), 412)

	if input_file and (converter.getFileExt(input_file.filename) in IMPORT_FILTER_MAP):
		input_base = converter.getFileBasename(input_file.filename)
		input_file.save(secure_filename(input_file.filename))

		output_file = "%s.%s" % (input_base, "pdf")

		try:
			converter.convertToPdf(input_file.filename, "pdf")
			return send_file(output_file)
		except Exception as e:
			error = str(e)
			app.logger.error(error)
			return make_response(jsonify({'error': error}), 500)
		finally:	
			os.remove(input_file.filename)	
			os.remove(output_file)
	error = "File is not of type .docx or .doc!"
	app.logger.error(error)
	return make_response(jsonify({'error': error}), 415)

if __name__ == "__main__":
	# Use this server for local development to get auto reloading
	if app.config['DEBUG']:
		app.run(host='0.0.0.0', port=5000)
	# Use this server for test/stage/production environment
	else:
		http_server = WSGIServer(('', 5000), app)
		http_server.serve_forever()