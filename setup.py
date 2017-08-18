from setuptools import setup
setup(
  name = 'DocxTemplateManager',
  packages = ['DocxTemplateManager'], # this must be the same as the name above
  version = '0.1.3',
  description = 'A library for managing templated Microsoft Word Documents with imports from Excel.',
  author = 'Brian Carlson',
  author_email = 'briancarlson6174@gmail.com',
  url = 'https://github.com/TheRobotCarlson/DocxTemplateManager', # use the URL to the github repo
  download_url = 'https://github.com/TheRobotCarlson/DocxTemplateManager/archive/0.1.tar.gz', # 
  install_requires=['DocxMerge', 'docxtpl', 'openpyxl'],
  keywords = ['docx', 'merge', 'word', 'excel'], # arbitrary keywords
)