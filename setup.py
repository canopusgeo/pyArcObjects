from setuptools import setup, find_packages

config = {
    'name' : 'pyArcObjects',
    'version' : "1.2.0",
    'description' : 'Utility Functions for using fine grained ArcObjects in Python and GP script Tools.',
    'author' : 'Canopus Geoinformatics Ltd',
    'author_email' : 'info@canopusgeo.com',
    'url' : 'https://github.com/canopusgeo/pyArcObjects',
    'packages' : find_packages(),
	'test_suite' : 'nose.collector',
    'install_requires' : ['nose'],
    'include_package_data' : True,
    'classifiers' : [
        'Environment :: Console',
        'Intended Audience :: Science/Research/GIS',
		'License :: GPLv2',
        'Programming Language :: Python :: 2.7',
        'Topic :: Scientific/Engineering/GIS',
        'Operating System :: Microsoft :: Windows'
    ]
    }
setup(**config)



