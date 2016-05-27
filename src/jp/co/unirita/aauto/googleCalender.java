package jp.co.unirita.aauto;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.ByteBuffer;
import java.text.DecimalFormat;
import java.util.Arrays;
import java.util.Calendar;
import java.util.List;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson.JacksonFactory;
import com.google.gdata.client.spreadsheet.CellQuery;
import com.google.gdata.client.spreadsheet.SpreadsheetQuery;
import com.google.gdata.client.spreadsheet.SpreadsheetService;
import com.google.gdata.client.spreadsheet.WorksheetQuery;
import com.google.gdata.data.spreadsheet.CellEntry;
import com.google.gdata.data.spreadsheet.CellFeed;
import com.google.gdata.data.spreadsheet.SpreadsheetEntry;
import com.google.gdata.data.spreadsheet.SpreadsheetFeed;
import com.google.gdata.data.spreadsheet.WorksheetEntry;
import com.google.gdata.data.spreadsheet.WorksheetFeed;
import com.google.gdata.util.ServiceException;

public class googleCalender {

    // アプリケーション名 (任意)
    private static final String APPLICATION_NAME = "SpreadsheetSearch/1.0";

    // サービス アカウント ID
    private static final String ACCOUNT_P12_ID = "spreadsheetsearch-1294@appspot.gserviceaccount.com";
    private static final File P12FILE = new File(
            "C:\\WORK\\MorningDuty\\SpreadsheetSearch-04f0a2a37c20.p12");

    // 認証スコープ
    private static final List<String> SCOPES = Arrays.asList(
            "https://docs.google.com/feeds",
            "https://spreadsheets.google.com/feeds");

    // Spreadsheet API URL
    private static final String SPREADSHEET_URL = "https://spreadsheets.google.com/feeds/spreadsheets/private/full";

    private static final URL SPREADSHEET_FEED_URL;
    // 固定
    private static final String memberListText = "C:\\WORK\\MorningDuty\\MemberList.txt";

    private static SpreadsheetService service = null;

    static {
        try {
            SPREADSHEET_FEED_URL = new URL(SPREADSHEET_URL);
        } catch (MalformedURLException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 認証処理
     *
     * @return
     * @throws Exception
     */
    private static Credential authorize() throws Exception {
        System.out.println("authorize in");

        HttpTransport httpTransport = GoogleNetHttpTransport
                .newTrustedTransport();
        JsonFactory jsonFactory = new JacksonFactory();

        GoogleCredential credential = new GoogleCredential.Builder()
                .setTransport(httpTransport).setJsonFactory(jsonFactory)
                .setServiceAccountId(ACCOUNT_P12_ID)
                .setServiceAccountPrivateKeyFromP12File(P12FILE)
                .setServiceAccountScopes(SCOPES).build();

        boolean ret = credential.refreshToken();
        // debug dump
        System.out.println("refreshToken:" + ret);

        // debug dump
        if (credential != null) {
            System.out.println("AccessToken:" + credential.getAccessToken());
        }

        System.out.println("authorize out");

        return credential;
    }

    /**
     * サービスの取得
     *
     * @return
     * @throws Exception
     */
    private static SpreadsheetService getService() throws Exception {
        System.out.println("service in");

        SpreadsheetService service = new SpreadsheetService(APPLICATION_NAME);
        service.setProtocolVersion(SpreadsheetService.Versions.V3);

        Credential credential = authorize();
        service.setOAuth2Credentials(credential);

        // debug dump
        System.out.println("Schema: " + service.getSchema().toString());
        System.out.println("Protocol: "
                + service.getProtocolVersion().getVersionString());
        System.out.println("ServiceVersion: " + service.getServiceVersion());

        System.out.println("service out");

        return service;
    }

    /**
     * スプレッドシート一覧(未使用)
     * @param service
     * @return
     * @throws Exception
     */
    private static List<SpreadsheetEntry> findAllSpreadsheets(
            SpreadsheetService service) throws Exception {
        System.out.println("findAllSpreadsheets in");

        SpreadsheetFeed feed = service.getFeed(SPREADSHEET_FEED_URL,
                SpreadsheetFeed.class);

        List<SpreadsheetEntry> spreadsheets = feed.getEntries();

        // debug dump
        for (SpreadsheetEntry spreadsheet : spreadsheets) {
            System.out.println("title: "
                    + spreadsheet.getTitle().getPlainText());
        }

        System.out.println("findAllSpreadsheets out");
        return spreadsheets;
    }

    /**
     * スプレッドシート名で検索
     * @param service
     * @param spreadsheetName
     * @return
     * @throws Exception
     */
    private static SpreadsheetEntry findSpreadsheetByName(SpreadsheetService service, String spreadsheetName) throws Exception {
        System.out.println("findSpreadsheetByName in");
        SpreadsheetQuery sheetQuery = new SpreadsheetQuery(SPREADSHEET_FEED_URL);
        sheetQuery.setTitleQuery(spreadsheetName);
        SpreadsheetFeed feed = service.query(sheetQuery, SpreadsheetFeed.class);
        SpreadsheetEntry ssEntry = null;
        if (feed.getEntries().size() > 0) {
          ssEntry = feed.getEntries().get(0);
        }
        System.out.println("findSpreadsheetByName out");
        return ssEntry;
      }

    /**
     * ワークシート名で検索
     * @param service スプレッドサービス
     * @param ssEntry スプレッドシートエンティティ
     * @param sheetName シート名
     * @return ワークシートエンティティ
     * @throws Exception
     */
    private static WorksheetEntry findWorksheetByName(SpreadsheetService service, SpreadsheetEntry ssEntry, String sheetName) throws Exception {
        System.out.println("findWorksheetByName in");
        WorksheetQuery worksheetQuery = new WorksheetQuery(ssEntry.getWorksheetFeedUrl());
        worksheetQuery.setTitleQuery(sheetName);
        WorksheetFeed feed = service.query(worksheetQuery, WorksheetFeed.class);
        WorksheetEntry wsEntry = null;
        if (feed.getEntries().size() > 0){
          wsEntry = feed.getEntries().get(0);
        }
        System.out.println("findWorksheetByName out");
        return wsEntry;
     }

    /**
     * 範囲指定クエリー
     * @param minrow
     * @param maxrow
     * @param mincol
     * @param maxcol
     * @return
     */
    private static String makeQuery(int minrow, int maxrow, int mincol, int maxcol) {
        String base = "?min-row=MINROW&max-row=MAXROW&min-col=MINCOL&max-col=MAXCOL";
        return base.replaceAll("MINROW", String.valueOf(minrow))
                .replaceAll("MAXROW", String.valueOf(maxrow))
                .replaceAll("MINCOL", String.valueOf(mincol))
                .replaceAll("MAXCOL", String.valueOf(maxcol));
    }

    /**
     *
     * @return
     * @throws ServiceException
     * @throws IOException
     */
    private static String getCellPlainText(WorksheetEntry worksheetEntry, String cellRange) throws IOException, ServiceException {

        CellQuery cellQuery = new CellQuery(worksheetEntry.getCellFeedUrl());

        cellQuery.setRange(cellRange);
        // 空セルも返すようにする
        cellQuery.setReturnEmpty(true);
        CellFeed cellFeed = service.query(cellQuery, CellFeed.class);

        CellEntry cellEntry;
        cellEntry = cellFeed.getEntries().get(0);

        return cellEntry.getPlainTextContent().toString();
    }

    /**
     * セル番号計算
     * @param worksheetEntry
     * @param cellRange 例："B2:B12"
     * @return
     * @throws IOException
     * @throws ServiceException
     */
    private static String getTargetCellNo(WorksheetEntry worksheetEntry, String cellRange) throws IOException, ServiceException {
        final String MONTH_PATTERN = "00";
        String targetCellNo = "A1";
        DecimalFormat format = new DecimalFormat(MONTH_PATTERN);
        Calendar nowCal = Calendar.getInstance();
        String month = String.valueOf(nowCal.get(Calendar.YEAR)) + "/"
                + format.format(nowCal.get(Calendar.MONTH) + 1);
        String row = nowCal.get(Calendar.DAY_OF_MONTH) + 3 + "";

        String[] cells = cellRange.split(":");
        int columnF = Integer.parseInt(String.valueOf(String.valueOf(cells[0].charAt(0)).getBytes("US-ASCII")[0]));
        int columnT = Integer.parseInt(String.valueOf(String.valueOf(cells[1].charAt(0)).getBytes("US-ASCII")[0]));
        String rowFT = cells[0].substring(1); // 列はA～Z固定

        for (int idx = columnF; idx <= columnT; idx++) {

            String cellNo = new String(int2Byte(idx)).trim() + rowFT;
            String monthCell = getCellPlainText(worksheetEntry, cellNo);

            if (monthCell != null && month.equals(monthCell)) {
                // 今月
                targetCellNo = new String(int2Byte(idx + 1)).trim() + row;
                break;
            }
        }
        return targetCellNo;
    }

    /**
     *
     * @param args スプレッド名、シート名、タイトル範囲
     *
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {

        System.out.println("main start");
        String member = "";

        service = getService();

        String spreadsheetName = "プロダクト開発部朝礼当番_20150501";
        String worksheetName = "2016年度上期";
        String monthRange = "B3:M3";

        spreadsheetName = args[0];
        worksheetName = args[1];
        monthRange = args[2];
        System.out.println("spreadsheetName: " + spreadsheetName);
        System.out.println("worksheetName: " + worksheetName);
        System.out.println("monthRange: " + monthRange);

        SpreadsheetEntry spreadsheetEntry = findSpreadsheetByName(service, spreadsheetName);
        WorksheetEntry worksheetEntry = findWorksheetByName(service, spreadsheetEntry, worksheetName);

        if (worksheetEntry != null) {

            String targetCellNo = getTargetCellNo(worksheetEntry, monthRange);
            member = getCellPlainText(worksheetEntry, targetCellNo);

            if (member.indexOf("→") >= 0 && member.split("→").length > 1) {
                member = member.split("→")[member.split("→").length - 1];
            }
            System.out.println("member : " + member);

        }

        int memberNo = Integer.parseInt(getMemberNo(trim(member)));

        System.out.println("main end");

        System.exit(memberNo);
    }

    /**
     * int型から byte 型配列 に変換
     * @param value
     * @return
     */
    private static byte[] int2Byte(int value) {
        return ByteBuffer.allocate(Integer.SIZE / Byte.SIZE).putInt(value)
                .array();
    }

    /**
     * 社員番号取得
     * @param memberName
     * @return
     */
    private static String getMemberNo(String memberName) {

        String memberNo = "10000";
        try {
            FileInputStream fileInputStream = new FileInputStream(memberListText);
            InputStreamReader inputStreamReader = new InputStreamReader(fileInputStream, "SJIS");
            BufferedReader br = new BufferedReader(inputStreamReader);

            String str = br.readLine();
            // 括弧を削除
            memberName = replaceBrackets(memberName);

            while (str != null) {
                // 括弧を削除
                str = replaceBrackets(str);
                if (str.indexOf(memberName) > 0) {

                    memberNo = str.split(",")[0];
                    break;
                }
                str = br.readLine();
            }

            System.out.println("memberNo = " + memberNo);
            br.close();
            inputStreamReader.close();
            fileInputStream.close();

        } catch (FileNotFoundException e) {
            System.out.println(e);
        } catch (IOException e) {
            System.out.println(e);
        }
        return memberNo;
    }

    private static final String SPACE_CHAR_HALF = " ";
    private static final String SPACE_CHAR_WIDE = "　";
    /**
     * 苗字前後の全角スペース及び半角スペースを除去
     * @param value
     * @return
     */
    private static String trim(String value) {

        return ltrim(rtrim(value));
    }

    /**
     * 文字列の前の全角スペース及び半角スペースを除去
     *
     * @param value
     * @return String
     */
    private static String ltrim(String value) {
      if (value == null || value.equals("")) return value;

      int pos = 0;

      for (int i = 0; i < value.length(); i++) {
        String s = String.valueOf(value.charAt(i));
        if (!s.equals(SPACE_CHAR_HALF) && !s.equals(SPACE_CHAR_WIDE)) {
            break;
        } else {
            pos = i + 1;
        }
      }

      if (pos > 0) {
        return value.substring(pos);
      } else {
        return value;
      }
    }

    /**
     * 文字列の後の全角スペースおよび半角スペースを除去
     *
     * @param value
     * @return String
     */
    private static String rtrim(String value) {
      if (value == null || value.equals("")) return value;

      int pos = 0;

      for (int i = value.length() - 1; i >= 0; i--) {
        String s = String.valueOf(value.charAt(i));
        if (!s.equals(SPACE_CHAR_HALF) && !s.equals(SPACE_CHAR_WIDE)) {
            break;
        } else {
            pos = i;
        }
      }

      if (pos > 0) {
        return value.substring(0, pos);
      } else {
        return value;
      }
    }

    /**
     * 名前中の全角、半角括弧を削除
     * @param name
     * @return
     */
    private static String replaceBrackets(String name) {

        name = name.replace("(", ""); // 半角
        name = name.replace(")", ""); // 半角
        name = name.replace("（", "");// 全角
        name = name.replace("）", "");// 全角

        return name;
    }
}