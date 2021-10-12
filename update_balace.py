import pyexcel
import psycopg2

connection = psycopg2.connect(  dbname=     'ddm0bfn9sojtr1', 
                                host=       'ec2-54-161-208-31.compute-1.amazonaws.com',
                                user=       'bposbinnkvtuhx', 
                                password=   '25b6386f513ee2e12e6ab8b2f2b69a12d656224a0ddb893d122273a5033932dc', 
                                port=       '5432')

if __name__ == "__main__":
    cursor = connection.cursor()
    cursor.execute("SELECT username, balance FROM clients")
    array = cursor.fetchall()
    data = []
    for item in array:
        if item[0] is not None:
            data.append(item)
    try:
        myfile = open("./balance.xlsx", "r+")
        pyexcel.save_as(array= data, dest_file_name="balance.xlsx")
        print("balance.xlsx обновлён")
    except IOError:
        print("Нельзя обновить файл, пока он открыт. Закрой balance.xlsx и попробуй снова.")
        
