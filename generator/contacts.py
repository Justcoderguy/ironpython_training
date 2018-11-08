from models.contact import Contact
import random
import string
import getopt
import sys
import os.path
import jsonpickle
__author__ = 'pzqa'

try:
    opts, args = getopt.getopt(sys.argv[1:], "n:f:", ["number of contacts", "file"])
except getopt.GetoptError as err:
    getopt.usage()
    sys.exit(2)

n = 5
f = "data/contact.json"

for o, a in opts:
    if o == "-n":
        n = int(a)
    elif o == "-f":
        f = a


def random_string(prefix, maxlen):
    symbols = string.ascii_letters + string.digits + string.punctuation + " "*10
    return prefix + "".join([random.choice(symbols) for i in range(random.randrange(maxlen))])


contact_empty = [Contact(f_name='', m_name='', l_name='', n_name='',
                         title='', company='', address_one='',
                         home='', mobile='', work='',
                         fax='', email_one='',
                         email_two='', email_three='',
                         homepage='', bday='-', bmonth='-', byear='',
                         aday='-', amonth='-', ayear='',
                         address_two='', phone_two='',
                         notes='')]
contact_random = [Contact(f_name=random_string("f_name", 10), m_name=random_string("m_name", 10),
                          l_name=random_string("l_name", 10), n_name=random_string("n_name", 10),
                          title=random_string("title", 10), company=random_string("company", 10),
                          address_one=random_string("address_one", 10),
                          home=random_string("home", 10), mobile=random_string("mobile", 10),
                          work=random_string("work", 10), fax=random_string("fax", 10),
                          email_one=random_string("email_one", 10),
                          email_two=random_string("email_two", 10),
                          email_three=random_string("email_three", 10),
                          homepage=random_string("homepage", 10), bday='-', bmonth='-', byear='',
                          aday='-', amonth='-', ayear='',
                          address_two=random_string("address_two", 10),
                          phone_two=random_string("phone_two", 12),
                          notes=random_string("notes", 50)) for i in range(n)]

test_data = contact_empty + contact_random

file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", f)

with open(file, "w") as out:
    jsonpickle.set_encoder_options("json", indent=2)
    out.write(jsonpickle.encode(test_data))
