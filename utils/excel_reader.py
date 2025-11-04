import pandas as pd

def read_excel_row_by_row(file_path):
    """Read Excel and return a list of rows as dictionaries"""
    try:
        df = pd.read_excel(file_path)
        
        # Clean column names
        df.columns = (
            df.columns
            .str.strip()
            .str.replace('\s+', ' ', regex=True)
        )
        
        # Return list of rows as dictionaries
        return [row.to_dict() for _, row in df.iterrows()]
    
    except Exception as e:
        raise Exception(f"Excel read error: {str(e)}")
