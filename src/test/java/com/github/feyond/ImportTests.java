package com.github.feyond;

import com.github.feyond.excel.ImportExcel;
import com.github.feyond.excel.OperateType;
import com.github.feyond.excel.annotation.ExcelField;
import com.github.feyond.excel.annotation.ExcelSheet;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

import java.io.IOException;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.List;

/**
 * @author chenfy
 * @create 2017-10-24 14:04
 **/
public class ImportTests {
    public static void main(String[] args) throws IOException {
//        System.out.println(new ImportExcel("X:\\home\\2.xlsx")
//                .getDataList(0, 2));
//
//        System.out.println(new ImportExcel("X:\\home\\2.xlsx")
//                .getDataListWithHeader(2, 1));

        final ImportExcel importExcel = new ImportExcel("/usr/local/var/www/qten/temp/b13506363b57eaa411d21eac9ec4261b.xls");
        List<String> headers = importExcel.getHeaders(0, 0);
        System.out.println(headers);
        final List<A> dataList = importExcel
                .getDataList(0, 1, true, A.class);
        System.out.println(dataList);
//        System.out.println(new ImportExcel("X:\\home\\2.xlsx").getDataList(3,2, A.class, 1));

    }

    @Data
    @NoArgsConstructor
    @ToString
    public static class A {
        /*
        *
        订单编号 买家公司名 买家会员名 卖家公司名 卖家会员名 货品总价(元) 运费(元) 涨价或折扣(元) 实付款(元) 订单状态 订单创建时间
        订单付款时间 发货方 收货人姓名 收货地址 邮编 联系电话 联系手机 货品标题 单价(元) 数量 单位 货号 型号 Offer ID SKU ID
        物料编号 单品货号 货品种类 买家留言 物流公司运单号

        * */

        @ExcelField(header = "订单编号")
        private String orderNo;
        @ExcelField(header = "买家公司名")
        private String buyerName;
        @ExcelField(header = "买家会员名")
        private String buyerId;

        @ExcelField(header = "卖家公司名")
        private String sellerName;
        @ExcelField(header = "卖家会员名")
        private String sellerId;

        @ExcelField(header = "货品总价(元)")
        private Double totalPrice;
        @ExcelField(header = "运费(元)")
        private Double deliveryFee;
        @ExcelField(header = "涨价或折扣(元)")
        private Double otherFee;
        @ExcelField(header = "实付款(元)")
        private Double realPay;

        @ExcelField(header = "订单状态")
        private String status;
        @ExcelField(header = "订单创建时间")
        private LocalDateTime createTime;
        @ExcelField(header = "订单付款时间")
        private LocalDateTime payTime;

        @ExcelField(header = "发货方")
        private String shipper;
        @ExcelField(header = "收货人姓名")
        private String receiverName;
        @ExcelField(header = "收货地址")
        private String receiverAddress;
        @ExcelField(header = "邮编")
        private String postcode;
        @ExcelField(header = "联系电话")
        private String phone;
        @ExcelField(header = "联系手机")
        private String telephone;


        @ExcelField(header = "货品标题")
        private String skuName;
        @ExcelField(header = "单价(元)")
        private Double skuPrice;
        @ExcelField(header = "数量")
        private Integer quantity;
        @ExcelField(header = "单位")
        private String unit;


        @ExcelField(header = "货号")
        private String itemCode;
        @ExcelField(header = "型号")
        private String itemType;
        @ExcelField(header = "Offer ID")
        private String itemId;
        @ExcelField(header = "SKU ID")
        private String skuId;
        @ExcelField(header = "物料编号")
        private String itemNum;
        @ExcelField(header = "单品货号")
        private String skuCode;
        @ExcelField(header = "货品种类")
        private Integer skuNum;

        @ExcelField(header = "买家留言")
        private String buyerMessage;
        @ExcelField(header = "物流公司运单号")
        private String waybillNumber;

        @ExcelField(header = "发票：购货单位名称")
        private String invoiceCompany;
        @ExcelField(header = "发票：纳税人识别号")
        private String invoiceTaxpayerId;
        @ExcelField(header = "发票：地址、电话")
        private String invoiceAddressInfo;
        @ExcelField(header = "发票：开户行及账号")
        private String invoiceBankInfo;
        @ExcelField(header = "发票收取地址")
        private String invoiceReceivingAddress;


        @ExcelField(header = "关联编号")
        private String outerNo;
    }
}
