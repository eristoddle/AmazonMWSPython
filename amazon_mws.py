#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# Basic interface to Amazon MWS
# stephanmil@gmail.com
# Original by
# richard@sitescraper.net
# Edited, extended and update for newer version of api

import re, time, webbrowser, urllib, urllib2, hashlib, hmac, base64
from pprint import pprint
from xml.dom import minidom

class MWSError(Exception):
    pass

class Bag(): 
    '''Wrapper class for DOM nodes'''
    pass

class AmazonXML():

    def __init__(self, merchant_id, message_type, order_ob = None):
        self.xmldoc = minidom.Document()
        self.message_type = message_type
        self._add_envelope_node()
        self._add_header_node(merchant_id)
        self._add_message_type_node()
        self._message_id = 0

    def _add_envelope_node(self):
        self.am_env = self.xmldoc.createElement("AmazonEnvelope")
        self.am_env.setAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance')
        self.am_env.setAttribute('xsi:noNamespaceSchemaLocation', 'amzn-envelope.xsd')
        self.xmldoc.appendChild(self.am_env)

    def _add_header_node(self,  merchant_id):
        header = self.xmldoc.createElement('Header')
        self.am_env.appendChild(header)
        docver = self.xmldoc.createElement('DocumentVersion')
        header.appendChild(docver)
        merchid = self.xmldoc.createElement('MerchantIdentifier')
        header.appendChild(merchid)
        doctxt = self.xmldoc.createTextNode('1.01')
        docver.appendChild(doctxt)
        merchtxt = self.xmldoc.createTextNode(merchant_id)
        merchid.appendChild(merchtxt)

    def _add_message_type_node(self):
        message_type_node = self.xmldoc.createElement('MessageType')
        txt_node = self.xmldoc.createTextNode(self.message_type)
        message_type_node.appendChild(txt_node)
        self.am_env.appendChild(message_type_node)

    def add_message_node(self):
        self.message_node = self.xmldoc.createElement('Message')
        self.message_node.appendChild(self._message_id_node())
        self.message_node.appendChild(self._add_message_item())
        self.am_env.appendChild(self.message_node)

    def _message_id_node(self):
        self._message_id += 1
        message_id_node = self.xmldoc.createElement('MessageID')
        txt_node = self.xmldoc.createTextNode(str(self._message_id))
        message_id_node.appendChild(txt_node)
        return message_id_node

    def _add_message_item(self):
        message_item = self.xmldoc.createElement(self.message_type)
        #TODO: This wraps fulfillment_node_content which wraps item
        return message_item

    def fulfillment_node_content(self, amazon_oid, carrier, ship_meth, trackno):
        am_oid_node = self.xmldoc.createElement('AmazonOrderID')
        oid_txt_node = self.xmldoc.createTextNode(str(amazon_oid))
        am_oid_node.appendChild(oid_txt_node)
        fulfill_date_node = self.xmldoc.createElement('FulfillmentDate')
        date_txt_node = self.xmldoc.createTextNode(time.strftime("%Y-%m-%dT%H:%M:%S", time.gmtime()))
        fulfill_date_node.appendChild(date_txt_node)
        fulfill_data_node = self.xmldoc.createElement('FulfillmentData')
        carrier_node = self.xmldoc.createElement('CarrierCode')
        carrier_txt_node = self.xmldoc.createTextNode(carrier)
        carrier_node.appendChild(carrier_txt_node)
        ship_meth_node = self.xmldoc.createElement('ShippingMethod')
        ship_meth_txt_node = self.xmldoc.createTextNode(ship_meth)
        ship_meth_node.appendChild(ship_meth_txt_node)
        trackno_node = self.xmldoc.createElement('ShipperTrackingNumber')
        trackno_txt_node = self.xmldoc.createTextNode(str(trackno))
        trackno_node.appendChild(trackno_txt_node)

    def fulfillment_node_item(self, amazon_item_code,  quantity):
        """
        There can be 0 to unlimited of these in the FulfillmentData node
        """
        item_node = self.xmldoc.createElement('Item')
        order_item_id_node = self.xmldoc.createElement('AmazonOrderItemCode')
        order_item_txt_node = self.xmldoc.createTextNode(str(amazon_item_code))
        order_item_id_node.appendChild(order_item_txt_node)
        item_node.appendChild(order_item_id_node)
        order_item_qty_node = self.xmldoc.createElement('Quantity')
        order_item_qty_txt_node = self.xmldoc.createElement(str(quantity))
        order_item_qty_node.appendChild(order_item_qty_txt_node )
        item_node.appendChild(order_item_qty_node)
        return item_node

    def return_pretty_xml(self):
        return self.xmldoc.toprettyxml(indent="    ")

class MWS:
    def __init__(self, access_key, secret_key, merchant_id, marketplace_id, domain='https://mws.amazonaws.com', method = 'POST', user_agent='App/Version (Language=Python)'):
        self.access_key = access_key
        self.secret_key = secret_key
        self.merchant_id = merchant_id
        self.marketplace_id = marketplace_id
        self.domain = domain
        self.method = method
        self.user_agent = user_agent

    def make_request(self, request_data):
        """Make request to Amazon MWS API with these parameters
        """
        data = {
            'AWSAccessKeyId': self.access_key,
            'SellerId': self.merchant_id,
            'MarketplaceId.Id.1': self.marketplace_id,
            'SignatureMethod': 'HmacSHA256',
            'SignatureVersion': '2',
            'Timestamp': self.get_timestamp(),
            'Version': '2011-01-01'
        }
        data.update(request_data)
        keys = data.keys()
        keys.sort()
        values = map(data.get, keys)
        request_description = urllib.urlencode(zip(keys, values))
        request_description = request_description.replace('+', " ")
        signature = self.calc_signature(request_description)
        request = '%s?%s&Signature=%s' % (self.domain, request_description, urllib.quote(signature))
        if self.method == 'GET':
            try:
                xml = urllib2.urlopen(urllib2.Request(request, headers={'User-Agent': self.user_agent})).read()
            except urllib2.URLError, e:
                print e.code
                xml = e.read()
        elif self.method == 'POST':
            try:
                postdata = request_description + '&Signature=' + urllib.quote(signature)
                xml = urllib2.urlopen(urllib2.Request(self.domain, data=postdata, headers={'User-Agent': self.user_agent, \
                                                                                           'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'})).read()
            except urllib2.URLError, e:
                print e.code
                xml = e.read()
        if 'ErrorResponse' in xml:
            raise MWSError(xml)
        return self.xml_to_dict(xml)


    def calc_signature(self, request_description):
        """
        Calculate MWS signature to interface with Amazon
        """
        nohttpdom = self.domain.replace('https://', '')
        domparts = nohttpdom.split('/')
        domonly = domparts[0]
        dompath = self.domain.replace(domonly, '')
        sig_data = self.method + '\n' + domonly.lower() + '\n' + dompath.replace('https://', '') + '\n' + request_description
        return base64.b64encode(hmac.new(self.secret_key, sig_data, hashlib.sha256).digest())


    def xml_to_dict(self, xml):
        try:
            dom = minidom.parseString(xml)
        except:
            return xml
        else:
            return self.unmarshal(dom)


    def unmarshal(self, element, plugins=None, rc=None):
        """Return the Bag object with attributes populated using DOM element

        element: the root of the DOM element we are interested in
        plugins: callback functions to fine-tune the object structure
        rc: parent object, used in the recursive call

        This core function is inspired by Mark Pilgrim (f8dy@diveintomark.org)
        with some enhancement. Each node.tagName is evalued by plugins' callback
        functions:

        ﻿  if plugins['isBypassed'] is true:
        ﻿  ﻿  this elment is ignored
        ﻿  if plugins['isPivoted'] is true:
        ﻿  ﻿  this children of this elment is moved to grandparents
        ﻿  ﻿  this object is ignored.
        ﻿  if plugins['isCollective'] is true:
        ﻿  ﻿  this elment is mapped to []
        ﻿  if plugins['isCollected'] is true:
        ﻿  ﻿  this children of elment is appended to grandparent
        ﻿  ﻿  this object is ignored.
        """

        if(rc == None):
            rc = Bag()

        if(plugins == None):
            plugins = {}

        childElements = [e for e in element.childNodes if isinstance(e, minidom.Element)]

        if childElements:
            for child in childElements:
                key = child.tagName
                if hasattr(rc, key):
                    if type(getattr(rc, key)) <> type([]):
                        setattr(rc, key, [getattr(rc, key)])
                    setattr(rc, key, getattr(rc, key) + [self.unmarshal(child, plugins)])
                elif isinstance(child, minidom.Element):
                    if plugins.has_key('isPivoted') and plugins['isPivoted'](child.tagName):
                        self.unmarshal(child, plugins, rc)
                    elif plugins.has_key('isBypassed') and plugins['isBypassed'](child.tagName):
                        continue
                    elif plugins.has_key('isCollective') and plugins['isCollective'](child.tagName):
                        setattr(rc, key, self.unmarshal(child, plugins, wrappedIterator([])))
                    elif plugins.has_key('isCollected') and plugins['isCollected'](child.tagName):
                        rc.append(self.unmarshal(child, plugins))
                    else:
                        setattr(rc, key, self.unmarshal(child, plugins))
        else:
            rc = "".join([e.data for e in element.childNodes if isinstance(e, minidom.Text)])
        return rc


    def get_timestamp(self):
        return time.strftime("%Y-%m-%dT%H:%M:%S", time.gmtime())

    def list_orders(self, created_after,  status = "Unshipped"):
        plugins = {'isBypassed': ('Request',),}
        thebag = self.make_request({'Action' : 'ListOrders', 'CreatedAfter' : created_after})
        for order in thebag.ListOrdersResponse.ListOrdersResult.Orders.Order:
            if order.OrderStatus == status:
                try:
                    len(order.ShippingAddress.AddressLine2)
                except:
                    order.ShippingAddress.AddressLine2 = ''
                try:
                    len(order.ShippingAddress.Phone)
                except:
                    order.ShippingAddress.Phone = ''
                try:
                    len(order.ShippingAddress.AddressLine1)
                except:
                    order.ShippingAddress.AddressLine1 = ''
                yield {'amazon_oid' : order.AmazonOrderId, \
                       'order_total' : order.OrderTotal.Amount, \
                       'customer_name' : order.ShippingAddress.Name, \
                       'address_line_1' : order.ShippingAddress.AddressLine1, \
                       'address_line_2' : order.ShippingAddress.AddressLine2, \
                       'city' : order.ShippingAddress.City, \
                       'state' : order.ShippingAddress.StateOrRegion, \
                       'zip' : order.ShippingAddress.PostalCode, \
                       'country' : order.ShippingAddress.CountryCode, \
                       'phone': order.ShippingAddress.Phone, \
                       'purchase_date' : order.PurchaseDate}


    def list_order_items(self, order_id, throttle = 1.0):
        """
        This throttle prevents an Amazon throttle
        """
        time.sleep(throttle)
        plugins = {'isBypassed': ('Request',),}
        thebag = self.make_request({'Action' : 'ListOrderItems', 'AmazonOrderId' : order_id})
        orderlines = thebag.ListOrderItemsResponse.ListOrderItemsResult.OrderItems.OrderItem
        try:
            for item in orderlines:
                try:
                    yield {'amazon_iid' : item.OrderItemId, \
                           'quantity' : item.QuantityOrdered, \
                           'sku' : item.SellerSKU, \
                           'title' : item.Title, \
                           'price' : item.ItemPrice.Amount, \
                           'shipping' : item.ShippingPrice.Amount}
                except:
                    yield {'amazon_iid' : item.OrderItemId, \
                           'quantity' : item.QuantityOrdered, \
                           'sku' : item.SellerSKU, \
                           'title' : item.Title, \
                           'price' : item.ItemPrice.Amount, \
                           'shipping' : 0}
        except:
            try:
                yield {'amazon_iid' : orderlines.OrderItemId, \
                       'quantity' : orderlines.QuantityOrdered, \
                       'sku' : orderlines.SellerSKU, \
                       'title' : orderlines.Title, \
                       'price' : orderlines.ItemPrice.Amount, \
                       'shipping' : orderlines.ShippingPrice.Amount}
            except:
                yield {'amazon_iid' : orderlines.OrderItemId, \
                       'quantity' : orderlines.QuantityOrdered, \
                       'sku' : orderlines.SellerSKU, \
                       'title' : orderlines.Title, \
                       'price' : orderlines.ItemPrice.Amount, \
                       'shipping' : 0}


    def list_complete_orders(self, created_after, status = "Unshipped"):
        order_ids = []
        order_dict = {}
        for order in self.list_orders(created_after, status):
            order_dict[order['amazon_oid']] = order
        for oid, order in order_dict.iteritems():
            orderlines = [orderline for orderline in self.list_order_items(oid)]
            order_dict[oid]['orderlines'] = orderlines
        return order_dict

    def ship_order(self, ):
        pass


if __name__ == "__main__":
    #Retrieve Orders
    myMWS = MWS(AWS_ACCESS_KEY_ID, \
                AWS_SECRET_ACCESS_KEY, \
                MERCHANT_ID, \
                MARKETPLACE_ID,
                'https://mws.amazonservices.com/Orders/2011-01-01')
    print myMWS.list_complete_orders('2011-12-01', status="Unshipped")
