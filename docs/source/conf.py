#---------------------------------------------------------------------------------------------
# Copyright (c) Bentley Systems, Incorporated. All rights reserved.
# See COPYRIGHT.md in the repository root for full copyright notice
#---------------------------------------------------------------------------------------------
# -- Project information -----------------------------------------------------
import os
import sys
sys.path.insert(0, os.path.abspath('../..'))
project = 'openstaadpy'
copyright = 'Bentley Systems, Inc.'
author = 'Bentley Systems, Inc.'
release = '0.2'

# -- General configuration ---------------------------------------------------
extensions = [
    'sphinx.ext.autodoc',
    'sphinx.ext.autosummary',
    'sphinx.ext.napoleon',
    'sphinx.ext.viewcode',
    'sphinx_copybutton',
    'myst_parser',
]
autosummary_generate = True

autodoc_default_options = {
    'members': True,
    'undoc-members': True,
    'private-members': False,
    'inherited-members': True,
    'show-inheritance': True
}
autodoc_class_signature = 'separated'
autodoc_inherit_docstrings = False

templates_path = ['_templates']
exclude_patterns = []

copybutton_prompt_text = r">>> |\.\.\. "
copybutton_prompt_is_regexp = True

# -- Options for HTML output -------------------------------------------------
html_theme = "sphinx_book_theme"
html_logo = "_static/STAADPro_CONNECT_Icon_48x48-rounded.png"
html_favicon = "_static/STAADPro_CONNECT_Icon_32x32-rounded.png"
html_static_path = ['_static']
html_css_files = ['custom.css']


html_theme_options = {
    "use_issues_button": False,
    "use_download_button": False,
    "use_fullscreen_button": True,
    "show_navbar_depth": 3,
    "show_toc_level": 2,
}

