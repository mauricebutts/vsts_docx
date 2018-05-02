""" Author: Maurice Butts
Desc: This file acts as the API to the Datapack package. Use this to interface with the custom built commands.
"""
from vsts_datapack.DatapackDocx.DatapackDocx import _create_docx_table_from_query
from vsts_datapack.DatapackDocx.DatapackDocx import _add_datapack_hyperlink
from vsts_datapack.DatapackVsts.DatapackVsts import _datapack_query
from vsts_datapack.DatapackVsts.DatapackVsts import _datapack_item_count_query


def datapack_vsts_query(token, team_instance, query_id):
    """ Returns work_items, titles, priority, state, tag, customer_impact, and node_name from a query.
        This is a fairly straight forward function and will need to be edited in order to return something
        different. This is because what the vsts API returns us does not have intuitive naming so you'll have to
        go into this function and ping the API yourself.

    :param token        : string    : vsts token, should be in your config
    :param team_instance: string    : team name ex:('https://myteam.visualstudio.com/')
    :param query_id     : string    :  id of the query you'd like to use as found in the URL
    :return:              lists     : work_item_ids, returned_titles, returned_priority, returned_state, returned_tag
    """
    return _datapack_query(token, team_instance, query_id)


def datapack_vsts_item_count_query(token, team_instance, query_id, project):
    """ This func grabs the number of items and url of the query_id passed to it.

    :param token         : string       : VSTS token
    :param team_instance : string       : team instance ex.('https://myteam.visualstudio.com/')
    :param query_id      : string       : id of the query
    :param project       : string       : name of the project ex. ('https://myteam.visualstudio.com/myprojectname/')
    :return:               int, string  : number of query items, url to query
    """
    return _datapack_item_count_query(token, team_instance, query_id, project)


def create_docx_table_from_query(document, list_of_column_names, list_of_data_lists):
    """ Takes the document to act on and the list of column names to be converted into a table
        it then takes the list of lists to generate an entire table. Make sure that the list of
        of data lists matches order with the list of column names.

    :param document             : class         : Document object to be acted on.
    :param list_of_column_names : list          : WorkItem Ids
    :param list_of_data_lists   : list of lists : N number of lists to be turned into columns and created into the DatapackDocx table
    :return:                      class         : DatapackDocx.document.add_table
    """
    return _create_docx_table_from_query(document, list_of_column_names, list_of_data_lists)


def add_datapack_hyperlink(paragraph, url_text, query_id, token, team_instance, project):
    """

    :param paragraph    : class     : Paragraph object to be acted on
    :param url_text     : string    : Hyperlinked string that trails the query count
    :param query_id     : string    : Vsts query id
    :param token        : string    : Vsts auth token
    :param team_instance: string    : Vsts team instance
    :param project      : string    : Vsts project
    :return:  None
    """
    return _add_datapack_hyperlink(paragraph=paragraph, url_text=url_text, query_id=query_id,
                                   token=token, team_instance=team_instance, project=project)
