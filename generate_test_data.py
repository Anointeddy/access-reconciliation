import pandas as pd
import random

# Fake names
names = [
    "John Smith", "Jane Doe", "Michael Brown", "Sarah Johnson",
    "Chris Davis", "Amanda Wilson", "James Taylor", "Emily Martinez",
    "Daniel Anderson", "Laura Thomas", "Kevin Jackson", "Megan White",
    "Brian Harris", "Jessica Lewis", "Ryan Robinson", "Ashley Walker",
    "Matthew Hall", "Stephanie Allen", "Joshua Young", "Nicole King"
]

# Generate card numbers
all_cards = random.sample(range(100000, 999999), 8000)
ccure_cards = all_cards[:6000]
genetec_cards = all_cards[3000:]

# Build CCURE dataframe
ccure_data = {
    "Cardholder Name": [random.choice(names) for _ in range(6000)],
    "Card Number": ccure_cards,
    "Status": [random.choice(["Active", "Inactive"]) for _ in range(6000)]
}

# Build Genetec dataframe
genetec_data = {
    "Cardholder Name": [random.choice(names) for _ in range(5000)],
    "Card Number": genetec_cards,
    "Status": [random.choice(["Active", "Inactive"]) for _ in range(5000)]
}

# Save to Excel
pd.DataFrame(ccure_data).to_excel("data/ccure.xlsx", index=False)
pd.DataFrame(genetec_data).to_excel("data/genetec.xlsx", index=False)

print("=" * 40)
print("  TEST FILES CREATED SUCCESSFULLY")
print("=" * 40)
print(f"  CCURE   : 6,000 records")
print(f"  Genetec : 5,000 records")
print(f"  Files saved in data/ folder")
print("=" * 40)