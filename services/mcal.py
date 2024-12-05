import os
from config import  Confict


class McalML:
    """FileListChecker: Scan Folder and Generate File List"""

    config_file = Config.get_default_config_file(the_script_file=__file__)

    def __init__(self, config):
        self.auto_worker_name = self.__class__.__name__
        Worker.__init__(self, config)
        self.config = config
        self.MCAL_DOC_ROOT = os.environ["MCAL_DOC_ROOT"]
        self.MCAL_DEV_ROOT = os.environ["MCAL_DEV_ROOT"]