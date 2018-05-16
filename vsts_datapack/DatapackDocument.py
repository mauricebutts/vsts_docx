""" Author: Maurice Butts
Date: 5/16/2018
Desc: This file acts as the API to the Datapack package. Use this to interface with the custom built commands.
"""
from vsts_datapack.DatapackDocx.DatapackDocx import _create_docx_table_from_query
from vsts_datapack.DatapackDocx.DatapackDocx import _add_datapack_hyperlink
from vsts_datapack.DatapackVsts.DatapackVsts import _datapack_query
from vsts_datapack.DatapackVsts.DatapackVsts import _datapack_item_count_query
import docx
from docx.shared import Inches
import json

# TODO: Use a docx template to read from?
# TODO: Sort the tables


class DatapackDocument:

    def __init__(self, config_path=None):
        self.__create()

        # If no config is declared let's check current dir, if no config create a new one
        if not config_path:
            try:
                with open('config.json') as config_file:
                    self.config = json.load(config_file)
            except Exception:
                print('No config found, creating config in current dir')
                print('Please fill in the config file with your vsts creds')
                self.__create_config()
        else:
            print('Grabbing config from {}'.format(config_path))
            self.__grab_config(config_path)

    def __create(self):
        """ Start the document
        :return:
        """
        self.document = docx.Document()
        self.paragraph_header = self.document

    def __grab_config(self, config_path):
        """
        :param config_path:
        :return:
        """
        with open(config_path, 'w') as read_config:
            self.config = json.load(read_config)

    def __create_config(self):
        """
        :return:
        """
        config_format = {
           1: ['token', 'team_instance', 'project']
        }

        with open('config.json', 'w') as write_config:
            json.dump(config_format, write_config)

    def add_paragraph(self, text):
        self.__add_new_paragraph_and_set_header(text)

    def add_text(self, text):
        if self.paragraph_header == self.document:
            raise Exception("No text based object at head. Please start a paragraph or something.")
        self.paragraph_header.add_run(text)

    def add_linebreak(self):
        if self.paragraph_header == self.document:
            raise Exception("No text started, please add a paragraph.")
        self.__add_new_paragraph_and_set_header('')

    def page_break(self):
        self.document.add_page_break()
        self.__add_new_paragraph_and_set_header('')

    def write(self, path):
        self.document.save(path)

    def add_image(self, path, image_width):
        self.document.add_picture(path, width=Inches(image_width))

    def add_heading(self, text, heading_type):
        self.document.add_heading(text, heading_type)
        self.__add_new_paragraph_and_set_header('')

    def create_table_from_query(self, vsts_config_key, query_id):
        """ Create a table in the current location of the document based on the vsts work item query.
        """
        data_list_of_lists = self._datapack_vsts_query(vsts_config_key, query_id)
        column_names = ['work item ids', 'returned titles', 'returned priority', 'returned state', 'returned tag']
        self._datapack_sort(data_list_of_lists)
        self.create_docx_table(list_of_column_names=column_names, list_of_data_lists=data_list_of_lists)

    def add_datapack_hyperlink(self, url_text, vsts_config_key, query_id):
        """
        :param paragraph    : class     : Paragraph object to be acted on
        :param url_text     : string    : Hyperlinked string that trails the query count
        :param query_id     : string    : Vsts query id
        :param token        : string    : Vsts auth token
        :param team_instance: string    : Vsts team instance
        :param project      : string    : Vsts project
        :return:  None
        """
        vsts_creds = self.__grab_vsts_creds(str(vsts_config_key))
        return _add_datapack_hyperlink(paragraph=self.paragraph_header, url_text=url_text, query_id=query_id,
                                       token=vsts_creds[0], team_instance=vsts_creds[1], project=vsts_creds[2])

    def _datapack_vsts_query(self, vsts_config_key, query_id):
        """ Returns work_items, titles, priority, state, tag, customer_impact, and node_name from a query.
            This is a fairly straight forward function and will need to be edited in order to return something
            different. This is because what the vsts API returns us does not have intuitive naming so you'll have to
            go into this function and ping the API yourself.
        :param token        : string    : vsts token, should be in your config
        :param team_instance: string    : team name ex:('https://myteam.visualstudio.com/')
        :param query_id     : string    :  id of the query you'd like to use as found in the URL
        :return:              lists     : work_item_ids, returned_titles, returned_priority, returned_state, returned_tag
        """
        vsts_creds = self.__grab_vsts_creds(str(vsts_config_key))
        return _datapack_query(vsts_creds[0], vsts_creds[1], query_id)

    def _datapack_vsts_item_count_query(self, vsts_config_key, query_id):
        """ This func grabs the number of items and url of the query_id passed to it.
        :param token         : string       : VSTS token
        :param team_instance : string       : team instance ex.('https://myteam.visualstudio.com/')
        :param query_id      : string       : id of the query
        :param project       : string       : name of the project ex. ('https://myteam.visualstudio.com/myprojectname/')
        :return:               int, string  : number of query items, url to query
        """
        vsts_creds = self.__grab_vsts_creds(str(vsts_config_key))
        return _datapack_item_count_query(vsts_creds[0], vsts_creds[1], query_id, vsts_creds[2])

    def create_docx_table(self, list_of_column_names, list_of_data_lists):
        """ Takes the document to act on and the list of column names to be converted into a table
            it then takes the list of lists to generate an entire table. Make sure that the list of
            of data lists matches order with the list of column names.
        :param document             : class         : Document object to be acted on.
        :param list_of_column_names : list          : WorkItem Ids
        :param list_of_data_lists   : list of lists : N number of lists to be turned into columns and created into the DatapackDocx table
        :return:                      class         : DatapackDocx.document.add_table
        """
        return _create_docx_table_from_query(self.document, list_of_column_names, list_of_data_lists)

    def __grab_vsts_creds(self, config_key):
        return self.config[config_key]

    def __add_new_paragraph_and_set_header(self, text):
        paragraph = self.document.add_paragraph(text)
        self.paragraph_header = paragraph

    def _datapack_sort(self, list_of_lists):
        pass
