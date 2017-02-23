require 'spreadsheet'




# get all receipts from a store that have a certain field
# throw receipts in excel

def store(id)
  Store.find(id)
end

def receipts(store)
  store.receipts.where.not(client_phone: nil)
end

def status_receipts(receipts, status)
  receipts.select do |receipt|
    receipt.status == status
  end
end

def to_excel(store, receipts, spreadsheet)
  book = Spreadsheet::Workbook.new
  sheet1 = book.create_worksheet
  sheet1[0,0] = 'Store Name'
  sheet1[0,1] = 'Nombre Cliente'
  sheet1[0,2] = 'Telefono'
  sheet1[0,3] = 'Email'
  sheet1[0,3] = 'Status'
  sheet1[1,0] = store.name

  receipts.each_with_index do |rec, index |
  sheet1[index + 1 ,1] = rec.client_name
  sheet1[index + 1 ,2] = rec.client_phone
  sheet1[index + 1 ,3] = rec.client_email
  sheet1[index + 1 ,4] = rec.status
  end
  book.write("./app/workers/#{spreadsheet}.xlsx")
end

store = store(4544)
all_receipts  = receipts(store)
accepted = status_receipts(all_receipts, 'ACCEPTED')
expired = status_receipts(all_receipts, 'EXPIRED')
pending = status_receipts(all_receipts, 'PENDING')

full = accepted + expired + pending

to_excel(store, full, 'Monte')
