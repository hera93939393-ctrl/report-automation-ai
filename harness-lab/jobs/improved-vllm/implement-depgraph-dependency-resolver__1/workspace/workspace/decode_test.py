#!/usr/bin/env python3
"""Script to decode base64 manifests from the protected directory."""
import base64
import json

# Decode the manifests we saw
manifests = [
    ("flask.manifest", "eyJuYW1lIjogImZsYXNrIiwgInZlcnNpb24iOiAiMi4zLjAiLCAiZGVwZW5kZW5jaWVzIjogeyJ3ZXJremV1ZyI6ICI+PTIuMC4wIn19"),
    ("requests.manifest", "eyJuYW1lIjogInJlcXVlc3RzIiwgInZlcnNpb24iOiAiMi4yOC4wIiwgImRlcGVuZGVuY2llcyI6IHt9fQ=="),
    ("werkzeug.manifest", "eyJuYW1lIjogIndlcmt6ZXVnIiwgInZlcnNpb24iOiAiMi4yLjAiLCAiZGVwZW5kZW5jaWVzIjogeyt9fQ==")
]

for filename, encoded in manifests:
    decoded = base64.b64decode(encoded).decode('utf-8')
    json_obj = json.loads(decoded)
    print(f"{filename}: {json.dumps(json_obj, indent=2)}")
    print()
