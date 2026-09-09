import json
import hashlib

# Read the blockchain ledger JSON file
with open('blockchain_ledger.json', 'r') as f:
    content = f.read()

# Parse the JSON
data = json.loads(content)

# The ledger is a JSON string within an object, extract the transactions
transactions = json.loads(data['output'])

# The actual transactions content is in data['output']
full_data = json.loads(data['output'])
transactions = full_data['ledger']

print(f"Total transactions: {len(transactions)}")

# Verify each transaction
for i, tx in enumerate(transactions):
    tx_id = tx['transaction_id']
    sender = tx['sender']
    receiver = tx['receiver']
    amount = tx['amount']
    timestamp = tx['timestamp']
    previous_hash = tx['previous_hash']
    current_hash = tx['current_hash']
    
    # Calculate the expected current hash
    input_data = f"{tx_id}{sender}{receiver}{amount}{timestamp}{previous_hash}"
    expected_hash = hashlib.sha256(input_data.encode()).hexdigest()
    
    link_status = "LINK"  # previous_hash matches previous tx's current_hash
    
    if i > 0:
        if tx['previous_hash'] == transactions[i-1]['current_hash']:
            link_status = "OK"
        else:
            link_status = "LINK_BROKEN"
    
    if expected_hash == current_hash:
        current_status = "OK"
    else:
        current_status = "HASH_BROKEN"
    
    # Check for issues
    issues = []
    if link_status != "OK":
        issues.append(('previous_hash_chain', previous_hash))
    if current_status != "OK":
        issues.append(('current_hash', current_hash))
    
    if issues:
        print(f"Transaction {tx_id} ({i}):")
        for issue, bad_value in issues:
            print(f"  - Field: {issue}")
            print(f"    CORRUPT value: '{bad_value}'")
            print()

# Find the single corrupted transaction
# The task says there's only ONE corrupted transaction
