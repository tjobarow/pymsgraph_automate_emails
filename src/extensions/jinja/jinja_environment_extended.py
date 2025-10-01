#!/.venv-linux/bin/ python
# -*-coding:utf-8 -*-
"""
@File    :   jinja_environment_extended.py
@Time    :   2025-1-25 8:6:207
@Author  :   Thomas Obarowski
@Contact :   tjobarow@gmail.com
@User    :   tjobarow
@Version :   1.0
@License :   MIT License
@Desc    :   None
"""
from jinja2 import Environment, FileSystemLoader, Template
from loguru import logger


class JinjaFileSystemEnvironmentExtended(Environment):
    def __init__(self, template_file_path: str):
        """
        Initialize the extended Jinja2 Environment, which uses a FileSystemLoader.
        :param template_file_path:  Path to jinja2 template
        :type template_file_path: str
        """
        logger.debug("Initializing Jinja2 Environment Extended ⏳")
        template_path: str
        template_filename: str
        template_path, template_filename = self.split_template_path_and_filename(template_file_path)
        logger.debug(f"Template path: {template_path}")
        logger.debug(f"Template filename: {template_filename}")
        super().__init__(loader=FileSystemLoader(template_path))
        logger.debug("Initialized Jinja2 Environment with FileSystemLoader✅")
        self.template: Template = self.get_jinja2_template(template_filename)
        logger.debug("Loaded Jinja2 template ✅")
        logger.debug(f"Finished initializing Jinja2 Environment Extended ✅")

    @staticmethod
    def split_template_path_and_filename(template_file_path: str) -> tuple[str, str]:
        """
        Given the filepath to a jinja2 template, split the path into directory path
        and filename, resolve the directory relative path to absolute path, and return
        return the absolute path and template filename.
        :param template_file_path: Relative or absolute path to jinja2 template
        :type template_file_path: str
        :return: Template directory path and filename
        :rtype: tuple[str, str]
        """
        logger.trace(f"Entered function with template_file_path: {template_file_path}")
        # Handle both linux and windows paths, by matching on the slash type
        if "/" in template_file_path:
            path_delimiter: str = "/"
            logger.trace(f"Detected linux path type")
        else:
            path_delimiter: str = "\\"
            logger.trace(f"Detected windows path type")
        # Split the path by slash
        path_arr: list[str] = template_file_path.split(path_delimiter)
        logger.trace(f"Split path into array: {path_arr}")
        # The filename for the template is the last index of the split path
        # Pop the filename off of the list, this will remove the filename from the list
        template_filename: str = path_arr.pop(len(path_arr) - 1)
        logger.trace(f"Calculated filename to be {template_filename}")
        # Join the path together
        template_dir_path: str = f"{path_delimiter.join(path_arr)}{path_delimiter}"
        logger.trace(f"Calculated template directory to be {template_dir_path}")
        return template_dir_path, template_filename

    def get_jinja2_template(self, template_filename: str) -> Template:
        """
        Given the template filename, return the jinja2 template object.
        :param template_filename: Filename of jinja2 template
        :type template_filename: str
        :return: Jinja2 template object
        :rtype: Template
        """
        logger.trace(f"Entered function with template_filename: {template_filename}")
        return self.get_template(name=template_filename)
