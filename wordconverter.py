import uno
import subprocess
import time
import os
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException

from flask import jsonify, make_response
from doc_types import *

import logging

logger = logging.getLogger(__name__)

LIBREOFFICE_DEFAULT_PORT = 4000
LIBREOFFICE_DEFAULT_HOST = "localhost"

class DocumentConversionException(Exception):

    def get_message(self):
        return self._message

    def set_message(self, message):
        self._message = message

    message = property(get_message, set_message)

class DocumentConverter:

    def __init__(self, host=LIBREOFFICE_DEFAULT_HOST, port= LIBREOFFICE_DEFAULT_PORT):
        self.host = host
        self.port = port
        self.local_context = uno.getComponentContext()
        self.connect_str = "socket,host=%s,port=%s;urp;StarOffice.ComponentContext" % (LIBREOFFICE_DEFAULT_HOST, LIBREOFFICE_DEFAULT_PORT)
        self.resolver = self.local_context.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", self.local_context)
        self.runLibreInstance()

        try:
            self.context = self.resolver.resolve("uno:%s" % self.connect_str)
        except NoConnectException:
            raise DocumentConversionException("failed to connect to OpenOffice.org on port %s" % port)

        self.desktop = self.context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.context)

    def runLibreInstance(self):
        self.process = subprocess.Popen('soffice --headless --norestore --nofirststartwizard --accept="%s"' % self.connect_str, shell=True, stdin=None, stdout=None, stderr=None)
        logger.info("Starting LibreOffice...")
        time.sleep(2)

    def terminateProcess(self):
        self.process.kill()

    def convertToPdf(self, input_file, output_format):
        logger.info("Converting %s to a %s" % (input_file, output_format))
        input_url = self.toFileUrl(input_file)

        input_properties = {"Hidden": True}
        input_basename = self.getFileBasename(input_file)
        input_ext = self.getFileExt(input_file)

        output_file = "%s.%s" % (input_basename, output_format)
        output_url = self.toFileUrl(output_file)

        if input_ext in IMPORT_FILTER_MAP:
            input_properties.update(IMPORT_FILTER_MAP[input_ext])

        document = self.desktop.loadComponentFromURL(input_url, "_blank", 0, self.toProperties(input_properties))

        try:
            document.refresh()
        except Exception as e:
            logger.error("Error loading document: %s" % str(e))
            raise DocumentConversionException(e)

        doc_family = self.getFamily(document)

        output_ext = self.getFileExt(output_file)
        output_properties = self.getProperties(document, output_ext)

        try:
            document.storeToURL(output_url, self.toProperties(output_properties))
            document.close(True)
            logger.info("Document converted successfully!")
        except Exception as e:
            logger.error("Error converting document: %s" % str(e))
            raise DocumentConversionException(e)

    def getProperties(self, document, output_ext):
        family = self.getFamily(document)
        try:
            properties = EXPORT_FILTER_MAP[output_ext]
        except KeyError:
            raise DocumentConversionException("unknown output format: '%s'" % output_ext)
        try:
            return properties[family]
        except KeyError:
            raise DocumentConversionException("unsupported conversion: from '%s' to '%s'" % (family, output_ext))

    def getFamily(self, document):
        if document.supportsService("com.sun.star.text.GenericTextDocument"):
            return DOC_TEXT
        raise DocumentConversionException("unknown document family: %s" % document)

    def getFileExt(self, path):
        ext = os.path.splitext(path)[1]
        if ext is not None:
            return ext[1:].lower()

    def getFileBasename(self, path):
        name = os.path.splitext(path)[0]
        if name is not None:
            return name

    def toFileUrl(self, path):
        return uno.systemPathToFileUrl(os.path.abspath(path))

    def toProperties(self, propDict):
        properties = []
        for key in propDict:
            if type(propDict[key]) == dict:
                prop = PropertyValue(key, 0, uno.Any("[]com.sun.star.beans.PropertyValue",self.toProperties(propDict[key])), 0)
            else:
                prop = PropertyValue(key, 0, propDict[key], 0)
            properties.append(prop)
        return tuple(properties)
