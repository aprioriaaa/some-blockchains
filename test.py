import blockchain as bc
import datetime as date

# Read from Excel file
reader = bc.BlockchainReader()
blockchain = reader.read_blockchain_from_file("blockchain.xlsx")


# Process
blockchain.add_block("file5")
blockchain.add_block("file6")


# Save back to Excel file
blockchain.save_blockchain_to_file("blockchain.xlsx")

blockchain.print_blockchain()
