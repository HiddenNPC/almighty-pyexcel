from pymongo import MongoClient

def get_db_client(my_config):
    """
    build a mongodb connection using my_config.ini
    :param db_name:
    :return:
    """
    db_url = my_config['mongodb_connection']['db_url']
    db_port = int(my_config['mongodb_connection']['db_port'])
    db_username = my_config['mongodb_connection']['db_username']
    db_password = my_config['mongodb_connection']['db_password']
    db_if_auth = my_config['mongodb_connection']['db_if_auth']
    db_auth_dbname = my_config['mongodb_connection']['db_auth_dbname']
    db_name = my_config['mongodb_connection']['db_name']

    db_client = MongoClient(db_url, db_port)
    if db_if_auth:
        db_client[db_auth_dbname].authenticate(db_username, db_password, mechanism='SCRAM-SHA-1')
    db = db_client[db_name]
    return db
