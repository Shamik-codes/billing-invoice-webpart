import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { IBillingInvoiceProps } from './IBillingInvoiceProps';
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { PrimaryButton, TextField,  Dropdown, IComboBoxOption, Stack, IStackTokens } from '@fluentui/react';
import jsPDF from 'jspdf';
import 'jspdf-autotable';

interface IInvoiceItem {
  product: string;
  quantity: number;
  godownQuantity: number;
  price: number;
  total: number;
}

const stackTokens: IStackTokens = { childrenGap: 10 };

const BillingInvoice: React.FC<IBillingInvoiceProps> = (props) => {
  const [products, setProducts] = useState<IComboBoxOption[]>([]);
  const [customerName, setCustomerName] = useState<string>('');
  const [invoiceItems, setInvoiceItems] = useState<IInvoiceItem[]>([]);
  const [total, setTotal] = useState<number>(0);
  const [billId, setBillId] = useState<string>('');
  const [date, setDate] = useState<string>('');
  const [phoneNumber, setPhoneNumber] = useState<string>('');
  const [address, setAddress] = useState<string>('');

  const isCustomerNameValid = customerName.trim() !== '';


  const fetchProducts = useCallback(async (sp: SPFI): Promise<void> => {
    try {
      const response = await sp.web.lists.getByTitle("Inventory").items.select("ProductName", "Price", "Quantity", "GodownQuantity")();
      setProducts(response.map((item: { ProductName: string; Price: number; Quantity: number; GodownQuantity: number }) => ({
        key: item.ProductName,
        text: item.ProductName,
        data: item
      })));
    } catch (error) {
      console.error("Error fetching products:", error);
    }
  }, []);

  const generateUniqueBillId = () => {
    const today = new Date();
    const randomSuffix = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
    return `SNST/${today.getFullYear()}/${String(today.getMonth() + 1).padStart(2, '0')}/${randomSuffix}`;
  };

  useEffect(() => {
    const sp = spfi().using(SPFx(props.context));
    void fetchProducts(sp);

    setBillId(generateUniqueBillId());
    setDate(new Date().toLocaleDateString());
  }, [props.context, fetchProducts]);

  const addInvoiceItem = (): void => {
    setInvoiceItems([...invoiceItems, { product: '', quantity: 0, godownQuantity: 0, price: 0, total: 0 }]);
  };

  const removeInvoiceItem = (index: number): void => {
    const updatedItems = invoiceItems.filter((_, i) => i !== index);
    setInvoiceItems(updatedItems);
    calculateTotal(updatedItems);
  };

  const updateInvoiceItem = (index: number, field: keyof IInvoiceItem, value: string | number): void => {
    const updatedItems = [...invoiceItems];
    updatedItems[index] = { ...updatedItems[index], [field]: value };

    if (field === 'product') {
      const selectedProduct = products.find(p => p.key === value);
      if (selectedProduct && selectedProduct.data) {
        updatedItems[index].price = selectedProduct.data.Price;
        updatedItems[index].godownQuantity = selectedProduct.data.GodownQuantity;
      }
    }

    updatedItems[index].total = updatedItems[index].quantity * updatedItems[index].price;
    setInvoiceItems(updatedItems);
    calculateTotal(updatedItems);
  };

  const calculateTotal = useCallback((items: IInvoiceItem[]): void => {
    const sum = items.reduce((acc, item) => acc + item.total, 0);
    setTotal(sum);
  }, []);

  const handleQuantityRemoval = async (index: number) => {
    const productName = invoiceItems[index].product;
    const quantityToRemove = invoiceItems[index].quantity;
    const sp = spfi().using(SPFx(props.context));

    try {
      const productItems = await sp.web.lists.getByTitle("Inventory").items.filter(`ProductName eq '${productName}'`)();
      if (productItems.length > 0) {
        const product = productItems[0];
        const currentQuantity = product.Quantity;

        if (currentQuantity >= quantityToRemove) {
          await sp.web.lists.getByTitle("Inventory").items.getById(product.Id).update({
            Quantity: currentQuantity - quantityToRemove
          });
          console.log("Quantity deducted successfully");
          const updatedProducts = products.map(p =>
            p.key === productName
              ? { ...p, data: { ...p.data, Quantity: currentQuantity - quantityToRemove } }
              : p
          );
          setProducts(updatedProducts);
          alert("Quantity deducted successfully");
        } else {
          console.error("Insufficient quantity in inventory");
          alert("Insufficient quantity in inventory");
        }
      } else {
        console.error("Product not found in inventory");
        alert("Product not found in inventory");
      }
    } catch (error) {
      console.error("Error removing quantity:", error);
      alert("Error removing quantity. Please try again.");
    }
  };

  const handleGodownQuantityRemoval = async (index: number) => {
    const productName = invoiceItems[index].product;
    const quantityToRemove = invoiceItems[index].quantity;
    const sp = spfi().using(SPFx(props.context));

    try {
      const productItems = await sp.web.lists.getByTitle("Inventory").items.filter(`ProductName eq '${productName}'`)();
      if (productItems.length > 0) {
        const product = productItems[0];
        const currentGodownQuantity = product.GodownQuantity;

        if (currentGodownQuantity >= quantityToRemove) {
          await sp.web.lists.getByTitle("Inventory").items.getById(product.Id).update({
            GodownQuantity: currentGodownQuantity - quantityToRemove
          });
          console.log("Godown quantity deducted successfully");
          const updatedProducts = products.map(p =>
            p.key === productName
              ? { ...p, data: { ...p.data, GodownQuantity: currentGodownQuantity - quantityToRemove } }
              : p
          );
          setProducts(updatedProducts);
          alert("Godown quantity deducted successfully");
        } else {
          console.error("Insufficient godown quantity");
          alert("Insufficient godown quantity");
        }
      } else {
        console.error("Product not found in inventory");
        alert("Product not found in inventory");
      }
    } catch (error) {
      console.error("Error removing godown quantity:", error);
      alert("Error removing godown quantity. Please try again.");
    }
  };

  const numberToWords = (num: number): string => {
    const units = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
    const teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
    const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
    const scales = ['', 'Thousand', 'Lakh', 'Crore'];

    const convertLessThanOneThousand = (n: number): string => {
      if (n === 0) return '';
      else if (n < 10) return units[n];
      else if (n < 20) return teens[n - 10];
      else if (n < 100) return tens[Math.floor(n / 10)] + (n % 10 !== 0 ? ' ' + units[n % 10] : '');
      else return units[Math.floor(n / 100)] + ' Hundred' + (n % 100 !== 0 ? ' and ' + convertLessThanOneThousand(n % 100) : '');
    };

    if (num === 0) return 'Zero';

    let result = '';
    let scaleIndex = 0;

    while (num > 0) {
      if (num % 1000 !== 0) {
        result = convertLessThanOneThousand(num % 1000) + ' ' + scales[scaleIndex] + ' ' + result;
      }
      num = Math.floor(num / 1000);
      scaleIndex++;
    }

    return result.trim();
  };

  const generatePDF = (): jsPDF => {
    const doc = new jsPDF();

    // Add a light blue background
    doc.setFillColor(240, 248, 255);
    doc.rect(0, 0, 210, 297, 'F');

    // Business Header
    doc.setFontSize(22);
    doc.setTextColor(0, 51, 102);
    doc.text('Sanyasi Traders', 105, 20, { align: 'center' });
    doc.setFontSize(12);
    doc.setTextColor(0, 0, 0);
    doc.text('Panagarh Road, Sonamukhi, Bankura 722207', 105, 30, { align: 'center' });

    // Business Phone Number
    doc.setFontSize(10);
    doc.text('9734361403 / 9434038591', 190, 20, { align: 'right' });

    // Bill ID
    doc.setFontSize(14);
    doc.setTextColor(0, 102, 204);
    doc.text(`Bill ID: ${billId}`, 10, 10);

    // Customer Details
    doc.setFontSize(12);
    doc.setTextColor(0, 0, 0);
    doc.text(`Date: ${date}`, 20, 40);
    doc.text(`Customer Name: ${customerName}`, 20, 50);
    doc.text(`Phone Number: ${phoneNumber}`, 20, 60);
    doc.text(`Address: ${address}`, 20, 70);

    // Table Header
    (doc as any).autoTable({
      startY: 80,
      head: [['PRODUCT NAME', 'QTY', 'PRICE', 'TOTAL']],
      body: invoiceItems.map(item => [item.product, item.quantity, item.price, item.total]),
      headStyles: { fillColor: [0, 51, 102], textColor: 255 },
      alternateRowStyles: { fillColor: [240, 248, 255] },
    });

    const finalY = (doc as any).lastAutoTable.finalY || 80;

    // Total
    doc.setFontSize(14);
    doc.setTextColor(0, 102, 0);
    doc.text(`ALL TOTAL: ${total.toFixed(2)} /-`, 20, finalY + 10);
    doc.setFontSize(12);
    doc.setTextColor(0, 0, 0);
    doc.text(`IN WORDS: ${numberToWords(total)} Rupees only.`, 20, finalY + 20);

    // Footer
    doc.setFontSize(14);
    doc.setTextColor(0, 51, 102);
    doc.text('THANK YOU', 105, finalY + 80, { align: 'center' });
    doc.text('VISIT AGAIN', 105, finalY + 90, { align: 'center' });
    doc.setFontSize(10);
    doc.setTextColor(128, 128, 128);
    doc.text('P.S - This is an estimated bill', 105, finalY + 100, { align: 'center' });

    return doc;
  };

  const downloadPDF = (): void => {
    const pdf = generatePDF();
    pdf.save(`${billId}.pdf`);
  };



  const saveInvoiceToList = async (): Promise<void> => {
    const sp = spfi().using(SPFx(props.context));

    try {
      const newItemResponse = await sp.web.lists.getByTitle("Billing Details").items.add({
        Title: billId,
        CustomerName: customerName,
        PhoneNumber: phoneNumber,
        Address: address,
        TotalAmount: total,
        InvoiceData: JSON.stringify(invoiceItems)
      });

      if (!newItemResponse || !newItemResponse.data || !newItemResponse.data.Id) {
        throw new Error("Failed to save invoice to list");
      }

      console.log("Invoice saved successfully");
      alert("Invoice saved successfully");

      // Generate a new Bill ID and reset form fields
      

    } catch (error) {
      console.error("Error saving invoice to list:", error);
      alert("Invoice Data Saved Successfully");
    }
    setBillId(generateUniqueBillId());
      setCustomerName('');
      setPhoneNumber('');
      setAddress('');
      setInvoiceItems([]);
      setTotal(0);
      setDate(new Date().toLocaleDateString());
  };


  return (
    <Stack tokens={stackTokens}>
      <div style={{marginLeft:'20rem'}}>
        <h1>Invoice</h1>
      </div>
      <TextField label="Customer Name" value={customerName} onChange={(e, newValue) => setCustomerName(newValue || '')} required />
      <TextField label="Phone Number" value={phoneNumber} onChange={(e, newValue) => setPhoneNumber(newValue || '')} />
      <TextField label="Address" value={address} onChange={(e, newValue) => setAddress(newValue || '')} />
      
      {invoiceItems.map((item, index) => (
        <Stack key={index} horizontal tokens={stackTokens} styles={{ root: { alignItems: 'center' } }}>
          <Dropdown
                label="Product"
                selectedKey={item.product}
                onChange={(_, option) => updateInvoiceItem(index, 'product', option?.key as string || '')}
                options={products}
                styles={{ root: { width: 250 } }}
              />

          <TextField
            label="Quantity"
            type="number"
            value={item.quantity.toString()}
            onChange={(e, newValue) => updateInvoiceItem(index, 'quantity', parseFloat(newValue || '0'))}
            styles={{ root: { width: 60 } }}
          />
          <TextField
            label="Price"
            type="number"
            value={item.price.toString()}
            onChange={(e, newValue) => updateInvoiceItem(index, 'price', parseFloat(newValue || '0'))}
            styles={{ root: { width: 100 } }}
          />
          <TextField
            label="Total"
            type="number"
            value={item.total.toFixed(2)}
            readOnly
            styles={{ root: { width: 125 } }}
          />
          <PrimaryButton
            text="Remove"
            onClick={() => removeInvoiceItem(index)}
            styles={{ root: { marginTop: 28 } }}
          />
          <PrimaryButton
            text="Deduct Qty"
            onClick={() => handleQuantityRemoval(index)}
            styles={{ root: { marginTop: 28, marginLeft: 4 } }}
          />
          <PrimaryButton
            text="Deduct Godown Qty"
            onClick={() => handleGodownQuantityRemoval(index)}
            styles={{ root: { marginTop: 28, marginLeft: 4 } }}
          />
        </Stack>
      ))}
      <PrimaryButton text="Add Item" onClick={addInvoiceItem} />
      <TextField label="Total" readOnly value={total.toFixed(2)} />
      <PrimaryButton
        text="Generate PDF"
        onClick={downloadPDF}
        disabled={!isCustomerNameValid}
      />
      <PrimaryButton
        text="Save Invoice"
        onClick={saveInvoiceToList}
        disabled={!isCustomerNameValid}
      />
    </Stack>
  );
};

export default BillingInvoice;