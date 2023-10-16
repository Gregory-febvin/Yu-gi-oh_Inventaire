import json
import requests
import openpyxl

def update_exel(): 
    with open("yugioh_card_french.json", "r") as json_file:
        json_data = json.load(json_file)

    workbook = openpyxl.Workbook()

    # Target la feuille 2
    sheet = workbook.create_sheet("Sheet 2")

    #Entete de collone
    sheet.append([
        "Id carte",
        "Nom carte",
        "Type carte",
        "Atk carte",
        "Def carte",
        "Niveau carte",
        "Race carte",
        "Atribut carte",
        "Desc carte",
        "Nom extension",
        "Code extension",
        "Rarete carte",
        "Code rarete carte"
    ])

    # Iterate through the "data" list and write each card set to the Excel sheet
    for item in json_data["data"]:

        # Variable nécésaire, marche pas avec juste item[""]
        card_name = item["name"]
        atk = item.get("atk")
        defense = item.get("def")
        level = item.get("level")

        # Je me branle des carte skill de speed duel
        if item["type"] == "Skill Card":
            continue

        if ("atk" in item or "def" in item or "level" in item) and "card_sets" in item:
            # Carte monstre
            for card_set in item["card_sets"]:
                sheet.append([
                    item["id"],
                    item["name"],
                    item["type"],
                    atk,
                    defense,
                    level,
                    item["race"],
                    item["attribute"],
                    item["desc"],
                    card_set["set_name"],
                    card_set["set_code"],
                    card_set["set_rarity"],
                    card_set["set_rarity_code"]
                ])


            continue
        else:

            if "card_sets" in item:
                # Carte magie / piège 
                for card_set in item["card_sets"]:
                    sheet.append([
                        item["id"],
                        item["name"],
                        item["type"],
                        "",
                        "",
                        "",
                        item["race"],
                        "",
                        item["desc"],
                        card_set["set_name"],
                        card_set["set_code"],
                        card_set["set_rarity"],
                        card_set["set_rarity_code"]
                    ])
            else:
                # Carte sans carte set
                print(f"Card Name: {card_name}")

                sheet.append([
                    item["id"],
                    item["name"],
                    item["type"],
                    "",
                    "",
                    "",
                    item["race"],
                    "",
                    item["desc"],
                ])
                continue

    workbook.save("Inventaire.xlsx")

    print("Opération terminé")




# Get current version of api
db_version_url = "https://db.ygoprodeck.com/api/v7/checkDBVer.php"
response_db_version = requests.get(db_version_url)
db_version_data = response_db_version.json()
current_database_version = db_version_data[0]['database_version']

# Step 2: Load the database version of the stored data
try:
    with open('yugioh_card_french.json', 'r') as file:
        existing_data = json.load(file)
        existing_database_version = existing_data[0]['database_version']
except (FileNotFoundError, json.JSONDecodeError):
    existing_database_version = None

# Step 3: Compare version to know if a new version is available
if existing_database_version is None or current_database_version != existing_database_version:
    # Step 4: If database version different, get the new data for the cards
    card_info_url = f"https://db.ygoprodeck.com/api/v7/cardinfo.php?name=Fort&language=fr"
    response_card_info = requests.get(card_info_url)
    card_info_data = response_card_info.json()

    # Update the data for the cards and db version
    combined_data = [
        {
            "database_version": current_database_version,
            "last_update": db_version_data[0]['last_update']
        },
        card_info_data
    ]

    # Write data in the file
    with open('yugioh_card_french.json', 'w') as file:
        json.dump(combined_data, file, indent=4)

    # Update the exel sheet 2 with the new data
    update_exel()
else:
    print("Database version is up to date. No changes were made.")


