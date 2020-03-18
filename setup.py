# Copyright 2019 PyI40AAS Contributors
#
# Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with
# the License. You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an
# "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the
# specific language governing permissions and limitations under the License.

import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="pyecma376-2",
    version="0.2.2",
    author="Michael Thies",
    author_email="m.thies@plt.rwth-aachen.de",
    url="https://git.rwth-aachen.de/acplt/pyecma376-2",
    description="Library for reading and writing ECMA 376-2 (Open Packaging Conventions) files",
    long_description=long_description,
    long_description_content_type="text/markdown",
    packages=["pyecma376_2"],
    zip_safe=False,
    package_data={"pyecma376_2": ["py.typed"]},
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: Apache Software License",
        "Operating System :: OS Independent",
        "Development Status :: 3 - Alpha",
    ],
    python_requires='>=3.6',
    install_requires=[
        'lxml>=4.2,<5'
    ]
)
