import json
from collections import Counter

with open(r'd:\aaa umich2\lands\markers.json', 'r', encoding='utf-8') as f:
    markers = json.load(f)

print(f'Total markers: {len(markers)}')

primary_counts = Counter(m['primary'] for m in markers)
for k, v in primary_counts.most_common():
    print(f'  {k}: {v} ({v/len(markers)*100:.1f}%)')

print()
cat_counts = Counter(m['erected_by_category'] for m in markers)
for k, v in cat_counts.most_common():
    print(f'  {k}: {v}')

print()
no_coords = [m for m in markers if m['lat'] is None]
print(f'Markers without coords: {len(no_coords)}')
for m in no_coords:
    print(f"  ID {m['id']}: {m['title']}")

no_year = [m for m in markers if m['year'] is None]
print(f'Markers without year: {len(no_year)}')

print(f"\nFirst: id={markers[0]['id']}, title={markers[0]['title']}, lat={markers[0]['lat']}, lng={markers[0]['lng']}")
print(f"Last:  id={markers[-1]['id']}, title={markers[-1]['title']}, lat={markers[-1]['lat']}, lng={markers[-1]['lng']}")

indigenous = [m for m in markers if m['primary'] == 'Indigenous']
print(f'\nIndigenous markers ({len(indigenous)}):')
for m in indigenous:
    print(f"  ID {m['id']}: {m['title']}, lat={m['lat']}, lng={m['lng']}")
