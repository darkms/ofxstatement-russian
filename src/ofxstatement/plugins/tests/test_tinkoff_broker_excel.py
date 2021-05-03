from datetime import datetime

from ofxstatement.ui import UI
from ofxstatement.ofx import OfxWriter
from ofxstatement.plugins.tinkoff_broker_excel import TinkoffBrokerExcelPlugin
from .util import file_sample
from xml.dom.minidom import parseString

def test_parse_multiple_currencies_rub():
    expected_content = """<?xml version="1.0" ?>
<!--
OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:UTF-8
CHARSET:NONE
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE
-->
<OFX>
    <SIGNONMSGSRSV1>
        <SONRS>
            <STATUS>
                <CODE>0</CODE>
                <SEVERITY>INFO</SEVERITY>
            </STATUS>
            <DTSERVER>20210501000000</DTSERVER>
            <LANGUAGE>ENG</LANGUAGE>
        </SONRS>
    </SIGNONMSGSRSV1>
    <BANKMSGSRSV1>
        <STMTTRNRS>
            <TRNUID>0</TRNUID>
            <STATUS>
                <CODE>0</CODE>
                <SEVERITY>INFO</SEVERITY>
            </STATUS>
            <STMTRS>
                <CURDEF>RUB</CURDEF>
                <BANKACCTFROM>
                    <BANKID/>
                    <ACCTID>tinkoff broker RUB</ACCTID>
                    <ACCTTYPE>CHECKING</ACCTTYPE>
                </BANKACCTFROM>
                <BANKTRANLIST>
                    <DTSTART/>
                    <DTEND/>
                    <STMTTRN>
                        <TRNTYPE>XFER</TRNTYPE>
                        <DTPOSTED>20210101</DTPOSTED>
                        <TRNAMT>22827.00000</TRNAMT>
                        <FITID>5</FITID>
                        <MEMO>Продажа 300 USD по RUB 76,09. Сумма: RUB 22827, номер сделки: 5, номер поручения: 5</MEMO>
                        <BANKACCTTO>
                            <ACCTID>USD</ACCTID>
                            <ACCTTYPE>CHECKING</ACCTTYPE>
                        </BANKACCTTO>
                    </STMTTRN>
                    <STMTTRN>
                        <TRNTYPE>FEE</TRNTYPE>
                        <DTPOSTED>20210101</DTPOSTED>
                        <TRNAMT>-11.41000</TRNAMT>
                        <FITID>5-fees</FITID>
                        <MEMO>Комиссия за Продажа 300 USD по RUB 76,09. Сумма: RUB 22827, номер сделки: 5, номер поручения: 5</MEMO>
                    </STMTTRN>
                    <STMTTRN>
                        <TRNTYPE>DEBIT</TRNTYPE>
                        <DTPOSTED>20210101</DTPOSTED>
                        <TRNAMT>10000.00000</TRNAMT>
                        <FITID>0cd7660140d0742027c1545042f809a9ff1ce88a</FITID>
                        <MEMO>Пополнение счета None, зачислено 10000, списано 0, дата исполнения: 01.01.2021</MEMO>
                    </STMTTRN>
                    <STMTTRN>
                        <TRNTYPE>DEBIT</TRNTYPE>
                        <DTPOSTED>20210101</DTPOSTED>
                        <TRNAMT>1000.00000</TRNAMT>
                        <FITID>d7c2fc72457917658f3a553075fdfa36d6a2f0b2</FITID>
                        <MEMO>РЕПО None, зачислено 1000, списано 0, дата исполнения: 01.01.2021</MEMO>
                    </STMTTRN>
                    <STMTTRN>
                        <TRNTYPE>DEBIT</TRNTYPE>
                        <DTPOSTED>20210101</DTPOSTED>
                        <TRNAMT>-1000.00000</TRNAMT>
                        <FITID>aef3fb0c12ed520569d2ffb5429376ca4983d6ed</FITID>
                        <MEMO>РЕПО None, зачислено 0, списано 1000, дата исполнения: 01.01.2021</MEMO>
                    </STMTTRN>
                    <STMTTRN>
                        <TRNTYPE>SRVCHG</TRNTYPE>
                        <DTPOSTED>20210101</DTPOSTED>
                        <TRNAMT>-30.00000</TRNAMT>
                        <FITID>bbb4687d56793d924986477ab294a2b423bb6ee0</FITID>
                        <MEMO>Комиссия по тарифу None, зачислено 0, списано 30, дата исполнения: 01.01.2021</MEMO>
                    </STMTTRN>
                    <STMTTRN>
                        <TRNTYPE>OTHER</TRNTYPE>
                        <DTPOSTED>20210101</DTPOSTED>
                        <TRNAMT>-10.00000</TRNAMT>
                        <FITID>7013df65174a75346a4756a8bd6207905d84c30e</FITID>
                        <MEMO>Налог None, зачислено 0, списано 10, дата исполнения: 01.01.2021</MEMO>
                    </STMTTRN>
                    <STMTTRN>
                        <TRNTYPE>CREDIT</TRNTYPE>
                        <DTPOSTED>20210101</DTPOSTED>
                        <TRNAMT>-2000.00000</TRNAMT>
                        <FITID>5cc43d9a442f9938c43edeed6cbda711602f1f61</FITID>
                        <MEMO>Вывод средств None, зачислено 0, списано 2000, дата исполнения: 01.01.2021</MEMO>
                    </STMTTRN>
                </BANKTRANLIST>
                <LEDGERBAL>
                    <BALAMT/>
                    <DTASOF/>
                </LEDGERBAL>
            </STMTRS>
        </STMTTRNRS>
    </BANKMSGSRSV1>
    <SECLISTMSGSRSV1>
        <SECLIST>
            <STOCKINFO>
                <SECINFO>
                    <SECID>
                        <UNIQUEID>MVID.ME</UNIQUEID>
                        <UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>
                    </SECID>
                    <SECNAME>MVID.ME</SECNAME>
                    <TICKER>MVID.ME</TICKER>
                </SECINFO>
            </STOCKINFO>
        </SECLIST>
    </SECLISTMSGSRSV1>
    <INVSTMTMSGSRSV1>
        <INVSTMTTRNRS>
            <TRNUID>0</TRNUID>
            <STATUS>
                <CODE>0</CODE>
                <SEVERITY>INFO</SEVERITY>
            </STATUS>
            <INVSTMTRS>
                <DTASOF/>
                <CURDEF>RUB</CURDEF>
                <INVACCTFROM>
                    <BROKERID>Tinkoff Investments</BROKERID>
                    <ACCTID>tinkoff broker RUB</ACCTID>
                </INVACCTFROM>
                <INVTRANLIST>
                    <DTSTART/>
                    <DTEND/>
                    <SELLSTOCK>
                        <SELLTYPE>SELL</SELLTYPE>
                        <INVSELL>
                            <INVTRAN>
                                <FITID>4</FITID>
                                <DTTRADE>20210101</DTTRADE>
                                <MEMO>Продажа 14 М.видео (MVID) по RUB 854. Сумма: RUB 11956, комиссия RUB 5.98, номер сделки: 4, номер поручения: 4</MEMO>
                            </INVTRAN>
                            <SECID>
                                <UNIQUEID>MVID.ME</UNIQUEID>
                                <UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>
                            </SECID>
                            <SUBACCTSEC>OTHER</SUBACCTSEC>
                            <SUBACCTFUND>OTHER</SUBACCTFUND>
                            <FEES>5.98000</FEES>
                            <UNITPRICE>854.00000</UNITPRICE>
                            <UNITS>-14.00000</UNITS>
                            <TOTAL>11950.02000</TOTAL>
                        </INVSELL>
                    </SELLSTOCK>
                </INVTRANLIST>
            </INVSTMTRS>
        </INVSTMTTRNRS>
    </INVSTMTMSGSRSV1>
</OFX>
"""

    rub_ofx = get_ofx_for_currency('RUB')

    assert pretty_print_xml(rub_ofx) == expected_content


def test_parse_multiple_currencies_usd():
    expected_content = """<?xml version="1.0" ?>
<!--
OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:UTF-8
CHARSET:NONE
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE
-->
<OFX>
    <SIGNONMSGSRSV1>
        <SONRS>
            <STATUS>
                <CODE>0</CODE>
                <SEVERITY>INFO</SEVERITY>
            </STATUS>
            <DTSERVER>20210501000000</DTSERVER>
            <LANGUAGE>ENG</LANGUAGE>
        </SONRS>
    </SIGNONMSGSRSV1>
    <BANKMSGSRSV1>
        <STMTTRNRS>
            <TRNUID>0</TRNUID>
            <STATUS>
                <CODE>0</CODE>
                <SEVERITY>INFO</SEVERITY>
            </STATUS>
            <STMTRS>
                <CURDEF>USD</CURDEF>
                <BANKACCTFROM>
                    <BANKID/>
                    <ACCTID>tinkoff broker USD</ACCTID>
                    <ACCTTYPE>CHECKING</ACCTTYPE>
                </BANKACCTFROM>
                <BANKTRANLIST>
                    <DTSTART/>
                    <DTEND/>
                    <STMTTRN>
                        <TRNTYPE>DEBIT</TRNTYPE>
                        <DTPOSTED>20210101</DTPOSTED>
                        <TRNAMT>2000.00000</TRNAMT>
                        <FITID>a37644f0e4fd09823fd7efa782c67fd46194f7c0</FITID>
                        <MEMO>Пополнение счета None, зачислено 2000, списано 0, дата исполнения: 01.01.2021</MEMO>
                    </STMTTRN>
                    <STMTTRN>
                        <TRNTYPE>FEE</TRNTYPE>
                        <DTPOSTED>20210101</DTPOSTED>
                        <TRNAMT>-15.00000</TRNAMT>
                        <FITID>950a48edfae404ffe18e932c835fc448600c24b6</FITID>
                        <MEMO>Возмещение дохода по дивидендам - списание Microsoft/ 2 шт., зачислено 0, списано 15, дата исполнения: 01.01.2021</MEMO>
                    </STMTTRN>
                </BANKTRANLIST>
                <LEDGERBAL>
                    <BALAMT/>
                    <DTASOF/>
                </LEDGERBAL>
            </STMTRS>
        </STMTTRNRS>
    </BANKMSGSRSV1>
    <SECLISTMSGSRSV1>
        <SECLIST>
            <STOCKINFO>
                <SECINFO>
                    <SECID>
                        <UNIQUEID>AAPL</UNIQUEID>
                        <UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>
                    </SECID>
                    <SECNAME>AAPL</SECNAME>
                    <TICKER>AAPL</TICKER>
                </SECINFO>
                <SECINFO>
                    <SECID>
                        <UNIQUEID>MSFT</UNIQUEID>
                        <UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>
                    </SECID>
                    <SECNAME>MSFT</SECNAME>
                    <TICKER>MSFT</TICKER>
                </SECINFO>
                <SECINFO>
                    <SECID>
                        <UNIQUEID>GIS</UNIQUEID>
                        <UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>
                    </SECID>
                    <SECNAME>GIS</SECNAME>
                    <TICKER>GIS</TICKER>
                </SECINFO>
            </STOCKINFO>
        </SECLIST>
    </SECLISTMSGSRSV1>
    <INVSTMTMSGSRSV1>
        <INVSTMTTRNRS>
            <TRNUID>0</TRNUID>
            <STATUS>
                <CODE>0</CODE>
                <SEVERITY>INFO</SEVERITY>
            </STATUS>
            <INVSTMTRS>
                <DTASOF/>
                <CURDEF>USD</CURDEF>
                <INVACCTFROM>
                    <BROKERID>Tinkoff Investments</BROKERID>
                    <ACCTID>tinkoff broker USD</ACCTID>
                </INVACCTFROM>
                <INVTRANLIST>
                    <DTSTART/>
                    <DTEND/>
                    <BUYSTOCK>
                        <BUYTYPE>BUY</BUYTYPE>
                        <INVBUY>
                            <INVTRAN>
                                <FITID>1</FITID>
                                <DTTRADE>20210101</DTTRADE>
                                <MEMO>Покупка 3 Apple (AAPL) по USD 138,28. Сумма: USD 414,84, комиссия USD 1.24, номер сделки: 1, номер поручения: 1</MEMO>
                            </INVTRAN>
                            <SECID>
                                <UNIQUEID>AAPL</UNIQUEID>
                                <UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>
                            </SECID>
                            <SUBACCTSEC>OTHER</SUBACCTSEC>
                            <SUBACCTFUND>OTHER</SUBACCTFUND>
                            <FEES>1.24000</FEES>
                            <UNITPRICE>138.28000</UNITPRICE>
                            <UNITS>3.00000</UNITS>
                            <TOTAL>-416.08000</TOTAL>
                        </INVBUY>
                    </BUYSTOCK>
                    <BUYSTOCK>
                        <BUYTYPE>BUY</BUYTYPE>
                        <INVBUY>
                            <INVTRAN>
                                <FITID>2</FITID>
                                <DTTRADE>20210101</DTTRADE>
                                <MEMO>Покупка 2 Microsoft (MSFT) по USD 230,81. Сумма: USD 461,62, комиссия USD 1.38, номер сделки: 2, номер поручения: 2</MEMO>
                            </INVTRAN>
                            <SECID>
                                <UNIQUEID>MSFT</UNIQUEID>
                                <UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>
                            </SECID>
                            <SUBACCTSEC>OTHER</SUBACCTSEC>
                            <SUBACCTFUND>OTHER</SUBACCTFUND>
                            <FEES>1.38000</FEES>
                            <UNITPRICE>230.81000</UNITPRICE>
                            <UNITS>2.00000</UNITS>
                            <TOTAL>-463.00000</TOTAL>
                        </INVBUY>
                    </BUYSTOCK>
                    <SELLSTOCK>
                        <SELLTYPE>SELL</SELLTYPE>
                        <INVSELL>
                            <INVTRAN>
                                <FITID>3</FITID>
                                <DTTRADE>20210101</DTTRADE>
                                <MEMO>Продажа 5 Microsoft (MSFT) по USD 225,63. Сумма: USD 1128,15, комиссия USD 0.28, номер сделки: 3, номер поручения: 3</MEMO>
                            </INVTRAN>
                            <SECID>
                                <UNIQUEID>MSFT</UNIQUEID>
                                <UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>
                            </SECID>
                            <SUBACCTSEC>OTHER</SUBACCTSEC>
                            <SUBACCTFUND>OTHER</SUBACCTFUND>
                            <FEES>0.28000</FEES>
                            <UNITPRICE>225.63000</UNITPRICE>
                            <UNITS>-5.00000</UNITS>
                            <TOTAL>1127.87000</TOTAL>
                        </INVSELL>
                    </SELLSTOCK>
                    <BUYSTOCK>
                        <BUYTYPE>BUY</BUYTYPE>
                        <INVBUY>
                            <INVTRAN>
                                <FITID>6</FITID>
                                <DTTRADE>20210101</DTTRADE>
                                <MEMO>Покупка 1 General Mills-ао (GIS) по USD 60,41. Сумма: USD 60,41, комиссия USD 0.03, номер сделки: 6, номер поручения: 6</MEMO>
                            </INVTRAN>
                            <SECID>
                                <UNIQUEID>GIS</UNIQUEID>
                                <UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>
                            </SECID>
                            <SUBACCTSEC>OTHER</SUBACCTSEC>
                            <SUBACCTFUND>OTHER</SUBACCTFUND>
                            <FEES>0.03000</FEES>
                            <UNITPRICE>60.41000</UNITPRICE>
                            <UNITS>1.00000</UNITS>
                            <TOTAL>-60.44000</TOTAL>
                        </INVBUY>
                    </BUYSTOCK>
                    <INCOME>
                        <INCOMETYPE>DIV</INCOMETYPE>
                        <INVTRAN>
                            <FITID>a179b01b3b714304bc8055d0506cbf00dc2f3081</FITID>
                            <DTTRADE>20210101</DTTRADE>
                            <MEMO>Выплата дивидендов Apple/ 3 шт., зачислено 10, списано 0, дата исполнения: 01.01.2021</MEMO>
                        </INVTRAN>
                        <SECID>
                            <UNIQUEID>AAPL</UNIQUEID>
                            <UNIQUEIDTYPE>TICKER</UNIQUEIDTYPE>
                        </SECID>
                        <SUBACCTSEC>OTHER</SUBACCTSEC>
                        <SUBACCTFUND>OTHER</SUBACCTFUND>
                        <TOTAL>10.00000</TOTAL>
                    </INCOME>
                </INVTRANLIST>
            </INVSTMTRS>
        </INVSTMTTRNRS>
    </INVSTMTMSGSRSV1>
</OFX>
"""

    usd_ofx = get_ofx_for_currency('USD')

    assert pretty_print_xml(usd_ofx) == expected_content

def get_ofx_for_currency(currency: str):
    plugin = TinkoffBrokerExcelPlugin(UI(), {
        'currency': currency,
        'account': f'tinkoff broker {currency}',
    })
    statement = plugin.get_parser(file_sample('tinkoff-broker-report-sample.xlsx')).parse()

    assert statement is not None
    assert statement.account_id == f'tinkoff broker {currency}'
    assert statement.currency == currency

    writer = OfxWriter(statement)
    # Set the generation time so it is always predictable
    writer.genTime = datetime(2021, 5, 1, 0, 0, 0)
    return writer.toxml()

def pretty_print_xml(xmlstr: str):
    dom = parseString(xmlstr)
    return dom.toprettyxml().replace("\t", "    ").replace("<!-- ", "<!--")
