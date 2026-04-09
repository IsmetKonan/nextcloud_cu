import pandas as pd
import privatebinapi
import time

input_file = "user.xlsx"
output_file = "pwlinks.xlsx"
base_url = "https://pass.ece24.net"

users_df = pd.read_excel(input_file, dtype=str)
total_seconds = 10 * len(users_df)
print(f"Total users: {len(users_df)}")
print(f"Gonna take about {total_seconds} seconds (~{total_seconds/60:.1f} minutes) due to rate limits...")

def create_privatebin_link(password: str) -> str:
    while True:
        try:
            response = privatebinapi.send(
                base_url,
                text=password,
                expiration="1month",
                compression=None
            )
            return response["full_url"]

        except Exception as e:
            if "Please wait" in str(e):
                print("Rate limit → waiting 10s...")
                time.sleep(10)
            else:
                print(f"Error: {e}")
                return ""

output_data = []

for _, user in users_df.iterrows():
    userid = user.get("userid")
    password = user.get("password")

    if not userid or not password or password == "nan":
        continue

    pb_link = create_privatebin_link(password)

    output_data.append({
        "userid": userid,
        "pwlink": pb_link
    })

    print(f"Created link for {userid}")

    time.sleep(10)

output_df = pd.DataFrame(output_data)
output_df.to_excel(output_file, index=False)

print(f"\nDone. Saved to {output_file}")