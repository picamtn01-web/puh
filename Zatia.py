# Zatia.py

# Update: Changed year range management to only reflect years from 2025 to 2030.

# Code snippet around line 261
max_year = max(current_year + 1, 2030)

# Updated year list generation as follows:
years = [str(y) for y in range(2030, 2024, -1)]  # Limiting available years to 2025-2030
