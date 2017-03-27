from distutils.core import setup

setup(
 name='ac2rpt',
 author = "Qiang Lu modified Dennis Muhlestein's csv2ofx",
 version='0.1',
 packages=['ac2rpt'],
 package_dir={'ac2rpt':'src/ac2rpt'},
 scripts=['ac2rpt'],
 package_data={'ac2rpt':['*.xrc']}
)

