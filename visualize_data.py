import polars as pl



def main():
    data = pl.read_csv('clean_order2024.csv')
    print(data.head(5))
    
if __name__ == "__main__":
    main()
