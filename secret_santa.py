import uuid
import random

# sender - address, receiver - address
people_str = """peter
paul
mary"""
people = list(people_str.splitlines())
receiver_address = { i : uuid.uuid4() for i in people }

sender_address = dict.fromkeys(people)
receivers = list(receiver_address.keys())
random.shuffle(receivers)
for sender in people:
    receiver = receivers.pop()
    while sender == receiver:
        receivers.insert(0, receiver)
        receiver = receivers.pop()
    sender_address[sender] = receiver_address[receiver]

print('=== receivers addresses are all public')
for k,v in receiver_address.items():
    print(k, v)


print('=== senders keep their destination addresses secret')
c = ''
while c != 'q':
    print('please enter a sender name, or q (quit):')
    c = input()
    if sender_address.get(c) is not None:
        print(sender_address[c])
    elif c == 'q':
        break
    print('\n\n\n\n\n\n\n\n\n\n\n\n\n\n')


# to check the pairing results
# for sender, s_address in sender_address.items():
#     for receiver, r_address in receiver_address.items():
#         if r_address == s_address:
#             print(sender, 'will prepare gift and card for', receiver, 'address:', s_address)
#             break