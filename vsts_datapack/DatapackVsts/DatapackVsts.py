from vsts.vss_connection import VssConnection
from msrest.authentication import BasicAuthentication
import pprint


def _datapack_query(token, team_instance, query_id):

    credentials = BasicAuthentication('', token)
    connection = VssConnection(base_url=team_instance, creds=credentials)

    core_client = connection.get_client('DatapackVsts.work_item_tracking.v4_1.work_item_tracking_client.WorkItemTrackingClient')

    results2 = core_client.query_by_id(id=query_id,
                                        team_context=None,
                                        time_precision=None)

    # Set up list to grab work_item_ids
    work_item_ids = []

    work_item_list = results2.work_items
    for wi in work_item_list:
        # print(wi.id)
        work_item_ids.append(wi.id)

    # Set up final lists
    returned_work_items = []
    returned_titles = []
    returned_priority = []
    returned_state = []
    returned_tag = []
    returned_customer_impact = []
    returned_node_name = []

    for id in work_item_ids:
        returned_work_items.append(core_client.get_work_item(id).fields)

    # DEBUG PRINT
    pprint.pprint(returned_work_items)

    for work_item in returned_work_items:

        # Look for Title
        try:
            returned_titles.append(work_item['System.Title'])
        except KeyError:
            returned_titles.append('')

        # Look for Priority
        try:
            returned_priority.append(work_item['Microsoft.VSTS.Common.Priority'])
        except KeyError:
            returned_priority.append('')

        # Look for State
        try:
            returned_state.append(work_item['System.State'])
        except KeyError:
            returned_priority.append('')

        # Look for Tag
        try:
            returned_tag.append(work_item['System.Tags'])
        except KeyError:
            returned_tag.append('')

    return work_item_ids, returned_titles, returned_priority, returned_state, returned_tag


def _datapack_item_count_query(token, team_instance, query_id, project):
    """ This func grabs the number of items and url of the query_id passed to it.

    :param token: VSTS token
    :param team_instance: team instance ex.('https://myteam.visualstudio.com/')
    :param query_id: id of the query
    :param project: name of the project ex. ('https://myteam.visualstudio.com/myprojectname/')
    :return: number of query items, url to query
    """

    credentials = BasicAuthentication('', token)
    connection = VssConnection(base_url=team_instance, creds=credentials)

    core_client = connection.get_client('DatapackVsts.work_item_tracking.v4_1.work_item_tracking_client.WorkItemTrackingClient')

    results = core_client.query_by_id(id=query_id,
                                        team_context=None,
                                        time_precision=None)

    url = team_instance + project + "/_queries?id=" + query_id + "&_a=query"

    return len(results.work_items), url
