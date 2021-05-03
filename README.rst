~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Russian banks plugins for ofxstatement
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

`ofxstatement`_ is a tool to convert proprietary bank statement to OFX format,
suitable for importing to GnuCash. Plugin for ofxstatement parses a
particular proprietary bank statement format and produces common data
structure, that is then formatted into an OFX file.

`ofxstatement-russian`_ provides some Russian banks plugins for ofxstatement.


.. _ofxstatement: https://github.com/kedder/ofxstatement
.. _ofxstatement-russian: https://github.com/gerasiov/ofxstatement-russian

Supported banks:

* Avangard Bank (https://avangard.ru) plugin 'avangard'
* Tinkoff Bank (https://tinkoff.ru) plugin 'tinkoff'
* Tinkoff Investments Broker (https://tinkoff.ru/invest) plugin 'tinkoff_broker_excel'
* SberBank (https://sbrf.ru) plugins 'sberbank_csv', 'sberbank_txt'
* AlfaBank (https://www.alfabank.ru) plugin 'alfabank'
* VTB (https://www.vtb.ru) plugin 'vtb'


Avangard
--------

CSV statements for credit, debit or current account.

Tinkoff
-------

CSV statement for credit, debit and saving account.

Tinkoff Investments Broker
-------

Excel statement for investments transactions.
Note: If you get an error related to missing 'style' - open & save the Excel file and try again.

SberBankCSV
--------

CSV statement (available via "request statement by e-mail as Excel sheet" function).

SberBankTxt
--------

Legacy TXT statement (available via "request statement by e-mail" function) for debit card.

AlfaBank
-------

CSV statement for current account.

VTB
-------

CSV statement for debit card.


Plugin configuration parameters
===============================

avangard
--------

bank
        Bank id
        (default is 'Avangard')

account
        Account id

currency
        Currency
        (default is 'RUB')

tinkoff
-------

bank
        Bank id
        (default is 'Tinkoff')

account
        Account id

currency
        Currency
        (if not set, will be extracted from the first record)

tinkoff_broker_excel
-------

broker
        Broker id
        (default is 'Tinkoff Investments')

account
        Account id
        (your currency-specific broker account id, ex: tinkoff broker usd)

currency
        Currency
        (only transactions nominated in selected currency would be extracted, other transactions would be skipped)

sberbank_txt
--------

bank
        Bank id
        (default is 'SberBank')

account
        Account id
        (if not set, will be extracted from the statement)

currency
        Currency
        (if not set, will be extracted from the statement)

alfabank
--------

bank
        Bank id
        (default is 'Alfabank')

account
        Account id
        (if not set, will be extracted from the statement)

currency
        Currency
        (if not set, will be extracted from the statement)

user_date
        if 'true' then transaction date will be set to the date when transaction is created (so called user date)
        rather then record date. User date is extracted in description if it is present there

file_encoding
        cp1251 by default. No need to change in regular usage (download statement, then convert),
        but could be handy in case of some file processing that involves encoding change

vtb
--------

bank
        Bank id
        (default is 'VTB')

user_date
        if 'true' then transaction date will be set to the date when transaction is created (so called user date)
        rather then record date.

Development
===========

Project is targeting python 3 (3.6 for sure as current widespread version) and
pytest is used for testing.

Development setup is simple:

1. Create virtual environment and activate it
.. code-block:: bash

    virtualenv .venv

    # activate it according to your OS specifics

2. Install dependencies. It will download everything you need to develop and write tests
.. code-block:: bash

    pip install -r requirements.txt

    python setup.py develop

3. Run tests using pytest
.. code-block:: bash

    pytest



Authors
=======
|  Copyright (c) 2013 Andrey Lebedev <andrey@lebedev.lt>
|  Copyright (c) 2016-2017 Alexander Gerasiov <gq@cs.msu.su>
|  Copyright (c) 2017 Dmitry Pavlov <zeldigas@gmail.com>
|

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License version 3 as
published by the Free Software Foundation.
