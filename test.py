exchange_dn = "//O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)/CN=RECIPIENTS/CN=3820A0C319C14FD8A95C114D4CF019EF-USHA.SHETTY"

# Split the Exchange DN by "/"
parts = exchange_dn.split("/")
cn_part = None

# Find the part that starts with "cn="
for part in parts:
    if part.startswith("cn="):
        cn_part = part
        break

# Extract the CN value after "cn="
if cn_part:
    cn_value = cn_part.split("cn=")[1]
    email_address = cn_value
    print("Extracted Email Address:", email_address)
else:
    print("Unable to extract email address from Exchange DN.")
#/O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)/CN=RECIPIENTS/CN=12C3E919CF464AF6ADE4D3B5AF5276A7-VRUSHALI.GH
#/O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)/CN=RECIPIENTS/CN=F790DE9F3698434BA5601C80EE79AD98-RAVIRAJ.JAD
#ravirajsinhj@skapsindia.com
#raviraj.jadeja@skapsindia.com