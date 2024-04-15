import hashlib
import datetime
import pandas as pd


class Block:
    def __init__(self, index, timestamp, data, previous_hash):
        self.index = index
        self.timestamp = timestamp
        self.data = data
        self.previous_hash = previous_hash
        self.hash = self.calculate_hash()

    def calculate_hash(self):
        sha = hashlib.sha256()
        sha.update(
            str(self.index).encode("utf-8")
            + str(self.timestamp).encode("utf-8")
            + str(self.data).encode("utf-8")
            + str(self.previous_hash).encode("utf-8")
        )
        return sha.hexdigest()


class Blockchain:
    def __init__(self):
        self.chain = [self.create_genesis_block()]

    def create_genesis_block(self):
        return Block(0, datetime.datetime.now(), "Genesis Block", "0")

    def get_latest_block(self):
        return self.chain[-1]

    def add_block(self, data):
        index = len(self.chain)
        timestamp = datetime.datetime.now()
        previous_hash = self.get_latest_block().hash
        new_block = Block(index, timestamp, data, previous_hash)
        self.chain.append(new_block)

    def is_chain_valid(self):
        for i in range(1, len(self.chain)):
            current_block = self.chain[i]
            previous_block = self.chain[i - 1]

            if current_block.hash != current_block.calculate_hash():
                return False

            if current_block.previous_hash != previous_block.hash:
                return False

        return True

    def save_blockchain_to_file(self, file_name):
        blockchain_data = []
        for block in self.chain:
            blockchain_data.append(
                {
                    "Index": block.index,
                    "Timestamp": block.timestamp,
                    "Data": block.data,
                    "Hash": block.hash,
                    "Previous Hash": block.previous_hash,
                }
            )

        df = pd.DataFrame(blockchain_data)
        # file_name = file_name.split(".")[0] + ".xlsx"
        writer = pd.ExcelWriter(file_name, engine="xlsxwriter")
        df.to_excel(writer, index=False)
        writer._save()

    def print_blockchain(self):
        for block in self.chain:
            print(f"Block Index: {block.index}")
            print(f"Timestamp: {block.timestamp}")
            print(f"Block Data: {block.data}")
            print(f"Block Hash: {block.hash}")
            print(f"Previous Block Hash: {block.previous_hash}")
            print()


class BlockchainReader:
    @staticmethod
    def read_blockchain_from_file(file_name):
        chain = Blockchain()
        df = pd.read_excel(file_name)

        if len(df) > 0:
            # Read the genesis block from the file
            genesis_block_data = df.iloc[0]
            genesis_block = Block(
                index=genesis_block_data["Index"],
                timestamp=genesis_block_data["Timestamp"],
                data=genesis_block_data["Data"],
                previous_hash=genesis_block_data["Previous Hash"],
            )
            genesis_block.hash = genesis_block_data["Hash"]

            chain.chain[0] = genesis_block  # Replace the generated genesis block

        # Read the remaining blocks from the file
        for _, row in df.iloc[1:].iterrows():
            index = row["Index"]
            timestamp = row["Timestamp"]
            data = row["Data"]
            hash_value = row["Hash"]
            previous_hash = row["Previous Hash"]
            block = Block(index, timestamp, data, previous_hash)
            block.hash = hash_value

            chain.chain.append(block)

        return chain
