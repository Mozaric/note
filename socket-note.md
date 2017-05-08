[參考資料](http://wmnlab.ee.ntu.edu.tw/nmlab/exp1_socket.html "台灣大學 電機工程系 網路與多媒體實驗 - Socket Programming")

[參考資料](https://zh.wikipedia.org/wiki/Berkeley%E5%A5%97%E6%8E%A5%E5%AD%97 "Berkeley通訊端")

↑↑↑想對 Socket 有基本了解可以點選上方網址↑↑↑，以下僅擷取部分重點。

# Socket TCP/IP 簡介

## Socket 是甚麼

從網路的角度來看， Socket 就是通訊連結的端點；從程式設計者的角度來看，Socket 提供了一個良好的**介面**，使程式設計者不需知道下層網路協定運作的細節便可以**撰寫網路通訊程式**。

## Sockets 的分類

Sockets 可分為下面兩類：

1. Datagram Sockets(connectionless)(**UDP**)

資料在 Datagram Sockets 間是利用 UDP 封包傳送，因此接收端 Socket **可能會收到次序錯誤的資料**，且其中部分資料亦可能會遺失。

2. Stream Sockets(connection-oriented)(**TCP**)

資料在 Stream Sockets 間是利用 TCP 封包來傳送，因此接收端 Socket 可以收到**順序無誤、無重覆、正確的資料**。此外 TCP 傳送時是採資料流的方式，因在傳送時所有資料會視情況被分割在數個 TCP 封包中。

## 主從式架構模型 (Client/Server model)

每個網路應用程式都有一個通訊端點，而通訊端點類型又可以分成**用戶端 (Client)** 、**伺服器端 (Server)** 。

用戶端若想與伺服器端的 Socket 建立連線(Association)，首先，兩個 Socket 必須是同一種，同為 UDP 或 TCP ，第二，用戶端 Socket 需要伺服器端的 Socket Name 才能識別出伺服器端的 Socket ，才能進一步請求結合。就 TCP/IP 而言，一個 Socket Name 包括了**IP位址**、**連結埠編號**、以及**協定**本身。

## Socket TCP類 Server/Client 建立結合與傳輸流程圖

![Socket Flowchart](http://www.itread01.com/uploads/images/20161016/1476619886-9596.jpg "Socket 流程圖")

[流程圖來源](http://www.itread01.com/articles/1476619887.html "Socket編程（簡易聊天室客戶端/服務器編寫、CocoaAsyncSocket）")

1. Server 端開啟 Socket 並設定 Local End Point(IP、Port) ，然後與 End Point 進行 Bind()
2. Server 端 Socket 使用 Listen() 來監聽已 Bind 的 End Point ，使用 Accept() 等待 Client 端 Socket 要求連線
3. Client 端開啟 Socket
4. Client 端 Socket 設定 Remote End Point ，然後使用 Connect() 對該 End Point 發出連接請求
5. Server 端 Socket 收到連接請求後，產生一新 Socket ，其會與 Client 端 Socket 建立連線，而原本的 Socket 繼續等待新 Client 端 Socket 要求連線
6. 建立連接後， Server 端 Socket 與 Client 端 Socket 互相使用 Write() 與 Read() 或是 Send() 與 Receive() 來傳輸資料。

## Socket 應用

[Socket編程（簡易聊天室客戶端/服務器編寫、CocoaAsyncSocket）](http://www.itread01.com/articles/1476619887.html "Socket編程（簡易聊天室客戶端/服務器編寫、CocoaAsyncSocket）")

## Socket in VB

簡介 VB.NET 內的 System.Net.Sockets 類別的幾個常用的屬性與方法

詳細內容介紹請至 [Socket 類別](https://msdn.microsoft.com/zh-tw/library/system.net.sockets.socket(v=vs.110).aspx "System.Net.Sockets")

`IPEndPoint` : End Point Class

`Socket` : Socket Class

`Bind()` : 用於 Server 端，使 Socket 與 End Point **建立關聯**

`Listen()` : 用於 Server 端，**監聽**已關聯的 End Point

`BeginAccept()` : 用於 Server 端，開啟一個 Thread 來**接受 Client 端 Socket 要求連線**

`EndAccept()` : 用於 Server 端，關閉 BeginAccept() 所開啟的 Thread 並取得執行結果

`BeginConnect()` : 用於 Client 端，開啟一個 Thread 來**向 Server 端發送連線請求**

`EndConnect()` : 用於 Client 端，關閉 BeginConnect() 所開啟的 Thread 並取得執行結果

`BeginReceive()` : 用於 Server/Client 端，開啟一個 Thread 來**傳送**資料

`EndReceive()` : 用於 Server/Client 端，關閉 BeginReceive() 所開啟的 Thread 並取得執行結果

`BeginSend()` : 用於 Server/Client 端，開啟一個 Thread 來**接收**資料

`EndSend()` : 用於 Server/Client 端，關閉 BeginSend() 所開啟的 Thread 並取得執行結果

`Connected As Boolean` : 取得 Socket 連線狀態

## 做做看

請參考 MSDN 的範例來寫兩個程式，一個 Server 端 Socket 程式，一個 Client 端 Socket 程式，兩程式可建立連線，並互相傳送訊息。

[Server 端範例](https://msdn.microsoft.com/zh-tw/library/fx6588te(v=vs.110).aspx "非同步伺服器通訊端範例")

[Client 端範例](https://msdn.microsoft.com/zh-tw/library/bew39x2a(v=vs.110).aspx "非同步用戶端通訊端範例")

在看範例之前，建議你要先了解：
>1. Socket TCP/IP類 Server/Client 建立結合與傳輸的**流程**
>2. Byte() 轉換成 String ，StringBuilder 類別
>3. String 轉換成 Byte() ，Encoding.ASCII.GetBytes() 方法
>4. Thread ，執行緒
>5. ManualResetEvent ，執行緒同步機制
>6. [Asynchronous Programming Model (非同步作業)](https://msdn.microsoft.com/zh-tw/library/ms228963(v=vs.110).aspx "非同步作業")
>7. Delegate **委派**

## 進階

實作以下功能：

>1. 當 Client 端中斷連線時， Server 端能偵測到
>2. 當 Server 端中斷連線時， Client 端能偵測到
>3. 偵測 Client 端 Connect 失敗
>4. Server 端對**多個** Client 端傳送訊息，亦即廣播
>5. Server 端對**指定的**一個 Client 端傳送訊息
>6. Server 端與 Client 端互相**傳送檔案**
>7. Server 端能取得目前連線中的 Client Socket 的數量與列表
