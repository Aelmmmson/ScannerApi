using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using Oracle.ManagedDataAccess.Client;
using System.Threading;
using System.Linq;
using Microsoft.AspNetCore.Cors;

namespace ScannerApi.Controllers
{
    [ApiController]
    [Route("api/scanner")]
    [EnableCors("AllowFrontend")]
    public class ScannerController : ControllerBase, IDisposable
    {
        private const int DEFAULT_STRING_BUFFER_SIZE = 4096;
        private const int MAX_IMAGE_SIZE_BYTES = 10 * 1024 * 1024;
        private readonly string imageSavePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ScannedImages");
        private string? g_strLogFileName;
        private string g_strAppPath = AppDomain.CurrentDomain.BaseDirectory;
        private int g_hLogFile = -1;
        private static bool m_bDeviceOpened = false;
        private static string m_strCurrentDeviceName = "";
        private string m_strOptions = new string('\0', DEFAULT_STRING_BUFFER_SIZE);
        private string m_strDocInfo = "";
        private readonly string connectionString = "Data Source=10.203.14.169:9534/USGL;User Id=XVSCAN;Password=pass1234;";
        private bool disposed = false;
        private readonly StreamWriter? logWriter;
        private readonly object logLock = new object();
        private static DocType m_nDocType = DocType.CHECK;

        private enum DocType
        {
            CHECK,
            MSR,
            INVALID
        }

        public ScannerController()
        {
            SetupLogging();
            string logPath = Path.Combine(g_strAppPath, "debug.log");
            try
            {
                if (!Directory.Exists(imageSavePath))
                {
                    Directory.CreateDirectory(imageSavePath);
                    LogMessage($"Constructor: Created image save directory at {imageSavePath}");
                }
                logWriter = new StreamWriter(logPath, append: true) { AutoFlush = true };
                LogMessage("SetupLogging: Log file initialized at " + logPath);
                LogMessage($"Constructor: Initial DocType={m_nDocType}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to initialize log file or directories: {ex.Message}");
                LogMessage($"Constructor: Failed to initialize log file or directories: {ex.Message}");
            }
        }

        private void SetupLogging()
        {
            g_strLogFileName = Path.Combine(g_strAppPath, "ExcellaLog.txt");
            g_hLogFile = CreateFile(g_strLogFileName, GENERIC_READ | GENERIC_WRITE,
                FILE_SHARE_READ | FILE_SHARE_WRITE, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);

            if (g_hLogFile > 0)
            {
                MTMICRSetLogFileHandle(g_hLogFile);
                LogMessage("SetupLogging: Log file handle set successfully");
            }
            else
            {
                LogMessage("SetupLogging: Failed to create log file handle");
            }
        }

        private void LogMessage(string message)
        {
            lock (logLock)
            {
                if (logWriter != null && !logWriter.BaseStream.CanWrite)
                {
                    Console.WriteLine($"LogMessage: StreamWriter unavailable, message: {message}");
                    return;
                }
                try
                {
                    logWriter?.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}");
                }
                catch (ObjectDisposedException)
                {
                    Console.WriteLine($"LogMessage: StreamWriter disposed, message: {message}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"LogMessage: Error writing log: {ex.Message}, message: {message}");
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    LogMessage("Dispose: Closing resources");
                }

                if (g_hLogFile > 0)
                {
                    CloseHandle(g_hLogFile);
                    g_hLogFile = -1;
                    LogMessage("Dispose: Closed log file handle");
                }

                if (m_bDeviceOpened)
                {
                    MTMICRCloseDevice(m_strCurrentDeviceName);
                    m_bDeviceOpened = false;
                    LogMessage("Dispose: Closed device connection");
                }

                if (disposing && logWriter != null)
                {
                    try
                    {
                        logWriter.Dispose();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Dispose: Error disposing logWriter: {ex.Message}");
                        LogMessage($"Dispose: Error disposing logWriter: {ex.Message}");
                    }
                }

                disposed = true;
            }
        }

        ~ScannerController()
        {
            Dispose(false);
        }

        [HttpPost("set-doctype/{docType}")]
        public IActionResult SetDocType(string docType)
        {
            try
            {
                if (Enum.TryParse(docType.ToUpper(), out DocType parsedDocType) && parsedDocType != DocType.INVALID)
                {
                    m_nDocType = parsedDocType;
                    LogMessage($"SetDocType: Document type set to {m_nDocType}");
                    return Ok(new { success = true, message = $"Document type set to {m_nDocType}" });
                }
                LogMessage($"SetDocType: Invalid document type: {docType}");
                return BadRequest(new { success = false, message = $"Invalid document type: {docType}" });
            }
            catch (Exception ex)
            {
                LogMessage($"SetDocType: Error: {ex.Message}, StackTrace: {ex.StackTrace}");
                return StatusCode(500, new { success = false, message = $"Error setting document type: {ex.Message}" });
            }
        }

        [HttpGet("status")]
        public IActionResult GetDeviceStatus()
        {
            try
            {
                string? deviceName = GetFirstDevice();
                if (string.IsNullOrEmpty(deviceName))
                {
                    LogMessage("GetDeviceStatus: No device found.");
                    return Ok(new { connected = false, message = "No device found" });
                }

                LogMessage($"GetDeviceStatus: Found device: {deviceName}, Current DocType={m_nDocType}");

                if (!m_bDeviceOpened || m_strCurrentDeviceName != deviceName)
                {
                    if (m_bDeviceOpened)
                    {
                        MTMICRCloseDevice(m_strCurrentDeviceName);
                        LogMessage($"GetDeviceStatus: Closed previous device: {m_strCurrentDeviceName}");
                        m_bDeviceOpened = false;
                    }

                    m_strCurrentDeviceName = deviceName;
                    int nRetOpen = MTMICROpenDevice(deviceName);
                    LogMessage($"GetDeviceStatus: MTMICROpenDevice returned {nRetOpen}");
                    if (nRetOpen != MICR_ST_OK)
                    {
                        MTMICRCloseDevice(deviceName);
                        LogMessage("GetDeviceStatus: Retried MTMICRCloseDevice");
                        nRetOpen = MTMICROpenDevice(deviceName);
                        LogMessage($"GetDeviceStatus: Retry MTMICROpenDevice returned {nRetOpen}");
                        if (nRetOpen != MICR_ST_OK)
                        {
                            return Ok(new { connected = false, message = $"Failed to open device, code: {nRetOpen}" });
                        }
                    }
                    m_bDeviceOpened = true;
                }

                int nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                StringBuilder response = new StringBuilder(DEFAULT_STRING_BUFFER_SIZE);
                int nRet = MTMICRQueryInfo(deviceName, "DeviceStatus", response, ref nResponseLength);
                LogMessage($"GetDeviceStatus: MTMICRQueryInfo attempt 1 returned {nRet}, Response: {response}, ResponseLength: {nResponseLength}");

                if (nRet != MICR_ST_OK || string.IsNullOrEmpty(response.ToString()))
                {
                    LogMessage("GetDeviceStatus: Retrying MTMICRQueryInfo");
                    nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                    response.Clear();
                    nRet = MTMICRQueryInfo(deviceName, "DeviceStatus", response, ref nResponseLength);
                    LogMessage($"GetDeviceStatus: MTMICRQueryInfo attempt 2 returned {nRet}, Response: {response}, ResponseLength: {nResponseLength}");
                }

                bool isConnected = nRet == MICR_ST_OK && !string.IsNullOrEmpty(response.ToString());
                LogMessage($"GetDeviceStatus: Device is {(isConnected ? "connected" : "not connected")}");

                return Ok(new { connected = isConnected, deviceName = deviceName, statusResponse = response.ToString() });
            }
            catch (Exception ex)
            {
                LogMessage($"GetDeviceStatus: Error: {ex.Message}, StackTrace: {ex.StackTrace}");
                return StatusCode(500, new { error = $"Error checking device status: {ex.Message}" });
            }
        }

        [HttpGet("devices")]
        public IActionResult GetDeviceList()
        {
            try
            {
                List<string> devices = new List<string>();
                const int maxRetries = 3;
                const int maxDevices = 20;

                for (int attempt = 1; attempt <= maxRetries; attempt++)
                {
                    devices.Clear();
                    for (byte nTotalDev = 1; nTotalDev <= maxDevices; nTotalDev++)
                    {
                        StringBuilder strDeviceName = new StringBuilder(256);
                        int nRetCode = MTMICRGetDevice(nTotalDev, strDeviceName);
                        LogMessage($"GetDeviceList: Attempt {attempt}, Index {nTotalDev}, MTMICRGetDevice returned {nRetCode}, DeviceName: {strDeviceName}");

                        if (nRetCode == MICR_ST_DEVICE_NOT_FOUND)
                        {
                            LogMessage($"GetDeviceList: Device not found at index {nTotalDev}, stopping enumeration");
                            break;
                        }

                        if (nRetCode == MICR_ST_OK && !string.IsNullOrEmpty(strDeviceName.ToString()))
                        {
                            string deviceName = strDeviceName.ToString();
                            if (!devices.Contains(deviceName))
                            {
                                devices.Add(deviceName);
                                LogMessage($"GetDeviceList: Added device: {deviceName}");
                            }
                        }
                    }

                    if (devices.Count > 0)
                    {
                        LogMessage($"GetDeviceList: Found {devices.Count} devices on attempt {attempt}: {string.Join(", ", devices)}");
                        break;
                    }

                    LogMessage($"GetDeviceList: No devices found on attempt {attempt}, retrying...");
                    Thread.Sleep(1000);
                }

                if (devices.Count == 0)
                {
                    LogMessage("GetDeviceList: No devices found after all retries.");
                    return Ok(new { devices = new string[0], message = "No devices found" });
                }

                LogMessage($"GetDeviceList: Final list: {string.Join(", ", devices)}");
                return Ok(new { devices = devices.ToArray(), message = $"{devices.Count} device(s) found" });
            }
            catch (Exception ex)
            {
                LogMessage($"GetDeviceList: Error: {ex.Message}, StackTrace: {ex.StackTrace}");
                return StatusCode(500, new { error = $"Error retrieving device list: {ex.Message}" });
            }
        }

        [HttpPost("connect")]
        public IActionResult ConnectDevice()
        {
            try
            {
                if (m_bDeviceOpened)
                {
                    LogMessage($"ConnectDevice: Device already opened: {m_strCurrentDeviceName}, DocType={m_nDocType}");
                    return Ok(new { success = true });
                }

                m_strCurrentDeviceName = GetFirstDevice() ?? "";
                if (string.IsNullOrEmpty(m_strCurrentDeviceName))
                {
                    LogMessage("ConnectDevice: No device found");
                    return BadRequest(new { success = false, message = "No device found" });
                }

                int nRet = MTMICROpenDevice(m_strCurrentDeviceName);
                LogMessage($"ConnectDevice: MTMICROpenDevice returned {nRet}");
                if (nRet == MICR_ST_OK)
                {
                    m_bDeviceOpened = true;
                    LogMessage($"ConnectDevice: Device opened successfully: {m_strCurrentDeviceName}, DocType={m_nDocType}");
                    return Ok(new { success = true });
                }
                else
                {
                    LogMessage($"ConnectDevice: Failed to open device, code: {nRet}");
                    return BadRequest(new { success = false, message = $"Failed to open device, code: {nRet}" });
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ConnectDevice: Error: {ex.Message}, StackTrace: {ex.StackTrace}");
                return StatusCode(500, new { success = false, message = $"Error connecting to device: {ex.Message}" });
            }
        }

        [HttpPost("connect/{deviceName}")]
        public IActionResult ConnectSpecificDevice(string deviceName)
        {
            try
            {
                if (m_bDeviceOpened)
                {
                    MTMICRCloseDevice(m_strCurrentDeviceName);
                    LogMessage($"ConnectSpecificDevice: Closed previous device: {m_strCurrentDeviceName}");
                    m_bDeviceOpened = false;
                }

                if (string.IsNullOrEmpty(deviceName))
                {
                    LogMessage("ConnectSpecificDevice: Invalid device name");
                    return BadRequest(new { success = false, message = "Device name is required" });
                }

                m_strCurrentDeviceName = deviceName;
                int nRet = MTMICROpenDevice(deviceName);
                LogMessage($"ConnectSpecificDevice: MTMICROpenDevice for {deviceName} returned {nRet}");
                if (nRet == MICR_ST_OK)
                {
                    m_bDeviceOpened = true;
                    LogMessage($"ConnectSpecificDevice: Device opened successfully: {deviceName}, DocType={m_nDocType}");
                    return Ok(new { success = true });
                }
                else
                {
                    LogMessage($"ConnectSpecificDevice: Failed to open device {deviceName}, code: {nRet}");
                    return BadRequest(new { success = false, message = $"Failed to open device {deviceName}, code: {nRet}" });
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ConnectSpecificDevice: Error: {ex.Message}, StackTrace: {ex.StackTrace}");
                return StatusCode(500, new { success = false, message = $"Error connecting to device {deviceName}: {ex.Message}" });
            }
        }

        [HttpPost("scan")]
        public IActionResult ScanVoucher()
        {
            try
            {
                LogMessage($"ScanVoucher: Starting scan with DocType={m_nDocType}, Device={m_strCurrentDeviceName}");
                if (!m_bDeviceOpened || string.IsNullOrEmpty(m_strCurrentDeviceName))
                {
                    m_strCurrentDeviceName = GetFirstDevice() ?? "";
                    if (string.IsNullOrEmpty(m_strCurrentDeviceName))
                    {
                        LogMessage("ScanVoucher: No device found");
                        return BadRequest(new { success = false, message = "No device found" });
                    }

                    int nRetOpen = MTMICROpenDevice(m_strCurrentDeviceName);
                    LogMessage($"ScanVoucher: MTMICROpenDevice returned {nRetOpen}");
                    if (nRetOpen != MICR_ST_OK)
                    {
                        MTMICRCloseDevice(m_strCurrentDeviceName);
                        LogMessage("ScanVoucher: Retried MTMICRCloseDevice");
                        nRetOpen = MTMICROpenDevice(m_strCurrentDeviceName);
                        LogMessage($"ScanVoucher: Retry MTMICROpenDevice returned {nRetOpen}");
                        if (nRetOpen != MICR_ST_OK)
                        {
                            LogMessage($"ScanVoucher: Failed to open device, code: {nRetOpen}");
                            return BadRequest(new { success = false, message = $"Failed to open device, code: {nRetOpen}" });
                        }
                    }
                    m_bDeviceOpened = true;
                    LogMessage($"ScanVoucher: Device opened successfully: {m_strCurrentDeviceName}, DocType={m_nDocType}");
                }

                // Reset device state to ensure clean configuration
                if (m_bDeviceOpened)
                {
                    MTMICRCloseDevice(m_strCurrentDeviceName);
                    LogMessage($"ScanVoucher: Closed device {m_strCurrentDeviceName} for reset");
                    m_bDeviceOpened = false;
                    int nRetOpen = MTMICROpenDevice(m_strCurrentDeviceName);
                    LogMessage($"ScanVoucher: Reopened device {m_strCurrentDeviceName}, returned {nRetOpen}");
                    if (nRetOpen != MICR_ST_OK)
                    {
                        LogMessage($"ScanVoucher: Failed to reopen device, code: {nRetOpen}");
                        return BadRequest(new { success = false, message = $"Failed to reopen device, code: {nRetOpen}" });
                    }
                    m_bDeviceOpened = true;
                }

                // Validate document type
                if (m_nDocType == DocType.INVALID)
                {
                    LogMessage("ScanVoucher: Invalid document type detected");
                    return BadRequest(new { success = false, message = "Invalid document type" });
                }

                int nRet = SetupOptions();
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"ScanVoucher: SetupOptions failed, code: {nRet}");
                    return BadRequest(new { success = false, message = $"Failed to setup options, code: {nRet}" });
                }

                StringBuilder strResponse = new StringBuilder(DEFAULT_STRING_BUFFER_SIZE);
                int nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                m_strDocInfo = "";
                const int maxRetries = 3;
                int attempt;

                for (attempt = 1; attempt <= maxRetries; attempt++)
                {
                    LogMessage($"ScanVoucher: Attempt {attempt} calling MTMICRProcessCheck with Options={m_strOptions}");
                    nRet = MTMICRProcessCheck(m_strCurrentDeviceName, m_strOptions, strResponse, ref nResponseLength);
                    LogMessage($"ScanVoucher: Attempt {attempt} MTMICRProcessCheck returned {nRet}, ResponseLength={nResponseLength}, DocInfo={strResponse}");

                    if (nRet == MICR_ST_OK)
                    {
                        m_strDocInfo = strResponse.ToString();
                        if (!string.IsNullOrEmpty(m_strDocInfo))
                        {
                            break;
                        }
                        LogMessage($"ScanVoucher: Attempt {attempt} received empty DocInfo, retrying...");
                    }
                    else
                    {
                        LogMessage($"ScanVoucher: Attempt {attempt} MTMICRProcessCheck failed with code {nRet}");
                    }

                    if (attempt < maxRetries)
                    {
                        Thread.Sleep(2000);
                        strResponse.Clear();
                        nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                    }
                }

                if (nRet != MICR_ST_OK)
                {
                    if (nRet == MICR_ST_INVALID_FEED_TYPE)
                    {
                        LogMessage("ScanVoucher: Invalid document feed type for scan");
                        return BadRequest(new { success = false, message = "Invalid document feed type for scan" });
                    }
                    LogMessage($"ScanVoucher: All attempts failed, last code: {nRet}");
                    return BadRequest(new { success = false, message = $"Process check failed with code {nRet}" });
                }

                if (string.IsNullOrEmpty(m_strDocInfo))
                {
                    LogMessage("ScanVoucher: Empty DocInfo received after all attempts");
                    return BadRequest(new { success = false, message = "No data captured from scan" });
                }

                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();
                nRet = MTMICRGetValue(m_strDocInfo, "CommandStatus", "ReturnCode", strResponse, ref nResponseLength);
                string strReturnCode = strResponse.ToString();
                int nReturnCode = int.TryParse(strReturnCode, out int code) ? code : -1;
                LogMessage($"ScanVoucher: CommandStatus ReturnCode: {nReturnCode}");

                if (nReturnCode == 0)
                {
                    var voucherData = ExtractVoucherData();
                    LogMessage($"ScanVoucher: Voucher data extracted: voucherNo={voucherData.voucherNo ?? "null"}, checkNumber={voucherData.checkNumber ?? "null"}, narration={voucherData.narration}");
                    return Ok(voucherData);
                }
                else
                {
                    LogMessage($"ScanVoucher: Process failed with ReturnCode: {nReturnCode}");
                    return BadRequest(new { success = false, message = $"Process failed with ReturnCode {nReturnCode}" });
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ScanVoucher: Error: {ex.Message}, StackTrace: {ex.StackTrace}");
                return StatusCode(500, new { success = false, message = $"Error scanning voucher: {ex.Message}" });
            }
        }

        [HttpPost("save")]
        public IActionResult SaveToDatabase([FromBody] VoucherData? voucherData)
        {
            LogMessage("SaveToDatabase: Received voucher data");
            if (voucherData == null)
            {
                LogMessage("SaveToDatabase: Invalid voucher data");
                return BadRequest(new { success = false, message = "Invalid voucher data" });
            }

            byte[]? frontImageBytes = null;
            byte[]? backImageBytes = null;

            try
            {
                LogMessage($"SaveToDatabase: Processing voucher {voucherData.voucherNo ?? "null"}, DocType={m_nDocType}");

                if (!string.IsNullOrEmpty(voucherData.frontImage) && m_nDocType == DocType.CHECK)
                {
                    if (!IsValidBase64(voucherData.frontImage))
                    {
                        LogMessage($"SaveToDatabase: Invalid Base64 for frontImage: {voucherData.frontImage.Substring(0, Math.Min(50, voucherData.frontImage.Length))}...");
                        return BadRequest(new { success = false, message = "Front image is not a valid Base64 string" });
                    }
                    try
                    {
                        frontImageBytes = Convert.FromBase64String(voucherData.frontImage);
                        LogMessage($"SaveToDatabase: Front image size: {frontImageBytes.Length} bytes");
                        if (frontImageBytes.Length > MAX_IMAGE_SIZE_BYTES)
                        {
                            LogMessage("SaveToDatabase: Front image too large");
                            return BadRequest(new { success = false, message = "Front image exceeds size limit (10MB)" });
                        }
                    }
                    catch (FormatException ex)
                    {
                        LogMessage($"SaveToDatabase: frontImage FormatException: {ex.Message}, StackTrace: {ex.StackTrace}");
                        return BadRequest(new { success = false, message = "Front image is not a valid Base64 string" });
                    }
                }
                else if (m_nDocType == DocType.CHECK)
                {
                    LogMessage("SaveToDatabase: No front image provided for CHECK");
                }

                if (!string.IsNullOrEmpty(voucherData.backImage) && m_nDocType == DocType.CHECK)
                {
                    if (!IsValidBase64(voucherData.backImage))
                    {
                        LogMessage($"SaveToDatabase: Invalid Base64 for backImage: {voucherData.backImage.Substring(0, Math.Min(50, voucherData.backImage.Length))}...");
                        return BadRequest(new { success = false, message = "Back image is not a valid Base64 string" });
                    }
                    try
                    {
                        backImageBytes = Convert.FromBase64String(voucherData.backImage);
                        LogMessage($"SaveToDatabase: Back image size: {backImageBytes.Length} bytes");
                        if (backImageBytes.Length > MAX_IMAGE_SIZE_BYTES)
                        {
                            LogMessage("SaveToDatabase: Back image too large");
                            return BadRequest(new { success = false, message = "Back image exceeds size limit (10MB)" });
                        }
                    }
                    catch (FormatException ex)
                    {
                        LogMessage($"SaveToDatabase: backImage FormatException: {ex.Message}, StackTrace: {ex.StackTrace}");
                        return BadRequest(new { success = false, message = "Back image is not a valid Base64 string" });
                    }
                }
                else if (m_nDocType == DocType.CHECK)
                {
                    LogMessage("SaveToDatabase: No back image provided for CHECK");
                }

                try
                {
                    LogMessage("SaveToDatabase: Attempting database connection");
                    using (var connection = new OracleConnection(connectionString))
                    {
                        connection.Open();
                        LogMessage("SaveToDatabase: Database connection opened");
                        string query = "INSERT INTO mbank_cheques (TRANS_ID, IMAGE1, IMAGE2, NARRATION) " +
                                      "VALUES (:transId, :image1, :image2, :narration)";
                        using (var command = new OracleCommand(query, connection))
                        {
                            command.Parameters.Add("transId", OracleDbType.Varchar2).Value = string.IsNullOrEmpty(voucherData.voucherNo) ? DBNull.Value : voucherData.voucherNo;
                            command.Parameters.Add("image1", OracleDbType.Blob).Value = m_nDocType == DocType.CHECK && frontImageBytes != null ? frontImageBytes : DBNull.Value;
                            command.Parameters.Add("image2", OracleDbType.Blob).Value = m_nDocType == DocType.CHECK && backImageBytes != null ? backImageBytes : DBNull.Value;
                            command.Parameters.Add("narration", OracleDbType.Varchar2).Value = string.IsNullOrEmpty(voucherData.narration) ? DBNull.Value : voucherData.narration;

                            LogMessage("SaveToDatabase: Executing query");
                            int rowsAffected = command.ExecuteNonQuery();
                            if (rowsAffected == 0)
                            {
                                LogMessage($"SaveToDatabase: No rows inserted for voucher {voucherData.voucherNo ?? "null"}");
                                return StatusCode(500, new { success = false, message = "Failed to insert voucher into database" });
                            }
                            LogMessage($"SaveToDatabase: Inserted {rowsAffected} row(s) for voucher {voucherData.voucherNo ?? "null"}");
                        }
                    }

                    LogMessage($"SaveToDatabase: Successfully saved voucher {voucherData.voucherNo ?? "null"} to database");
                    return Ok(new { success = true, message = $"Voucher {voucherData.voucherNo ?? "null"} saved to database" });
                }
                catch (OracleException ex)
                {
                    LogMessage($"SaveToDatabase: OracleException: {ex.Message}, ErrorCode: {ex.ErrorCode}, StackTrace: {ex.StackTrace}");
                    return StatusCode(500, new { success = false, message = $"Database error: {ex.Message}" });
                }
                catch (Exception ex)
                {
                    LogMessage($"SaveToDatabase: Error: {ex.Message}, StackTrace: {ex.StackTrace}");
                    return StatusCode(500, new { success = false, message = $"Error saving to database: {ex.Message}" });
                }
            }
            catch (Exception ex)
            {
                LogMessage($"SaveToDatabase: Unexpected error: {ex.Message}, StackTrace: {ex.StackTrace}");
                return StatusCode(500, new { success = false, message = $"Unexpected error: {ex.Message}" });
            }
        }

        [HttpGet("view/{transId}")]
        public IActionResult ViewVoucherData(string transId)
        {
            try
            {
                LogMessage($"ViewVoucherData: Fetching data for TRANS_ID={transId}, DocType={m_nDocType}");
                if (string.IsNullOrEmpty(transId))
                {
                    LogMessage("ViewVoucherData: Invalid transId");
                    return BadRequest(new { success = false, message = "Voucher number is required" });
                }

                using (var connection = new OracleConnection(connectionString))
                {
                    connection.Open();
                    LogMessage("ViewVoucherData: Database connection opened");
                    string query = "SELECT TRANS_ID, NARRATION, IMAGE1, IMAGE2 FROM mbank_cheques WHERE UPPER(TRANS_ID) = :transId";
                    using (var command = new OracleCommand(query, connection))
                    {
                        command.Parameters.Add("transId", OracleDbType.Varchar2).Value = transId.ToUpper().Trim();

                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                string? voucherNo = reader["TRANS_ID"] as string;
                                string? narration = reader["NARRATION"] as string;
                                byte[]? frontImageBytes = reader["IMAGE1"] as byte[];
                                byte[]? backImageBytes = reader["IMAGE2"] as byte[];

                                var voucherData = new VoucherData
                                {
                                    voucherNo = voucherNo ?? "",
                                    narration = narration ?? "",
                                    voucherType = "",
                                    micr = "",
                                    frontImage = frontImageBytes != null ? Convert.ToBase64String(frontImageBytes) : "",
                                    backImage = backImageBytes != null ? Convert.ToBase64String(backImageBytes) : "",
                                    frontImagePath = "",
                                    backImagePath = "",
                                    trackData1 = "",
                                    trackData2 = "",
                                    trackData3 = "",
                                    mpData = "",
                                    cardType = "",
                                    magnePrintStatus = "",
                                    track1Status = "",
                                    track2Status = "",
                                    track3Status = "",
                                    getScore = "",
                                    deviceSerialNumber = "",
                                    dukptSerialNumber = "",
                                    encryptedSessionId = "",
                                    encryptedTrack1 = "",
                                    encryptedTrack2 = "",
                                    encryptedTrack3 = "",
                                    checkNumber = "",
                                    accountNumber = "",
                                    routingNumber = "",
                                    bankCode = "",
                                    checkDate = "",
                                    amount = "",
                                    amountWords = "",
                                    accountHolder = "",
                                    signature = ""
                                };

                                LogMessage($"ViewVoucherData: Found voucher {transId}, Narration={narration ?? "null"}, FrontImageLen={frontImageBytes?.Length ?? 0}, BackImageLen={backImageBytes?.Length ?? 0}");
                                return Ok(new { success = true, data = voucherData });
                            }
                            else
                            {
                                LogMessage($"ViewVoucherData: No data found for voucher {transId}");
                                return Ok(new { success = false, message = $"No data found for voucher {transId}" });
                            }
                        }
                    }
                }
            }
            catch (OracleException ex)
            {
                LogMessage($"ViewVoucherData: OracleException: {ex.Message}, ErrorCode: {ex.ErrorCode}, StackTrace: {ex.StackTrace}");
                return StatusCode(500, new { success = false, message = $"Database error: {ex.Message}" });
            }
            catch (Exception ex)
            {
                LogMessage($"ViewVoucherData: Error: {ex.Message}, StackTrace: {ex.StackTrace}");
                return StatusCode(500, new { success = false, message = $"Error fetching voucher data: {ex.Message}" });
            }
        }

        private bool IsValidBase64(string base64String)
        {
            if (string.IsNullOrEmpty(base64String) || base64String.Length % 4 != 0)
            {
                LogMessage("IsValidBase64: Invalid length or empty string");
                return false;
            }

            try
            {
                Span<char> buffer = stackalloc char[base64String.Length];
                base64String.CopyTo(buffer);
                for (int i = 0; i < buffer.Length; i++)
                {
                    char c = buffer[i];
                    if (!(char.IsLetterOrDigit(c) || c == '+' || c == '/' || c == '='))
                    {
                        LogMessage($"IsValidBase64: Invalid character at position {i}: {c}");
                        return false;
                    }
                }
                int paddingCount = base64String.EndsWith("==") ? 2 : base64String.EndsWith("=") ? 1 : 0;
                if (paddingCount > 2)
                {
                    LogMessage("IsValidBase64: Invalid padding count");
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                LogMessage($"IsValidBase64: Error: {ex.Message}, StackTrace: {ex.StackTrace}");
                return false;
            }
        }

        private string? GetFirstDevice()
        {
            List<string> devices = new List<string>();
            const int maxRetries = 5;
            const int maxDevices = 20;

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                devices.Clear();
                for (byte nTotalDev = 1; nTotalDev <= maxDevices; nTotalDev++)
                {
                    StringBuilder strDeviceName = new StringBuilder(256);
                    int nRetCode = MTMICRGetDevice(nTotalDev, strDeviceName);
                    LogMessage($"GetFirstDevice: Attempt {attempt}, Index {nTotalDev}, MTMICRGetDevice returned {nRetCode}, DeviceName: {strDeviceName}");
                    if (nRetCode == MICR_ST_OK && !string.IsNullOrEmpty(strDeviceName.ToString()))
                    {
                        string deviceName = strDeviceName.ToString();
                        if (!devices.Contains(deviceName))
                        {
                            devices.Add(deviceName);
                            LogMessage($"GetFirstDevice: Added device: {deviceName}");
                        }
                    }
                    else if (nRetCode == MICR_ST_DEVICE_NOT_FOUND)
                    {
                        LogMessage($"GetFirstDevice: Device not found at index {nTotalDev}, stopping enumeration for attempt {attempt}");
                        break;
                    }
                }

                string? selectedDevice = devices.Find(d => d.Equals("STX.STX001", StringComparison.OrdinalIgnoreCase)) ?? devices.FirstOrDefault();
                if (selectedDevice != null)
                {
                    LogMessage($"GetFirstDevice: Selected device: {selectedDevice} on attempt {attempt}, DocType={m_nDocType}");
                    return selectedDevice;
                }

                LogMessage($"GetFirstDevice: No devices found on attempt {attempt}, retrying...");
                Thread.Sleep(2000);
            }

            LogMessage("GetFirstDevice: No devices found after all retries.");
            return null;
        }

        private int SetupOptions()
        {
            StringBuilder strOptions = new StringBuilder(DEFAULT_STRING_BUFFER_SIZE);
            strOptions.Capacity = DEFAULT_STRING_BUFFER_SIZE;
            int nActualLength = DEFAULT_STRING_BUFFER_SIZE;
            int nRet;

            if (m_nDocType == DocType.MSR)
            {
                nRet = MTMICRSetValue(strOptions, "ProcessOptions", "DocFeed", "MSR", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set DocFeed=MSR, code: {nRet}");
                    return nRet;
                }
                nRet = MTMICRSetValue(strOptions, "ProcessOptions", "DocFeedTimeout", "15000", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set DocFeedTimeout=15000, code: {nRet}");
                    return nRet;
                }
                nRet = MTMICRSetValue(strOptions, "ProcessOptions", "MSRFmt", "ISO", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set MSRFmt=ISO, code: {nRet}");
                    return nRet;
                }
            }
            else if (m_nDocType == DocType.CHECK)
            {
                nRet = MTMICRSetValue(strOptions, "Application", "Transfer", "HTTP", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set Transfer=HTTP, code: {nRet}");
                    return nRet;
                }

                nRet = MTMICRSetValue(strOptions, "Application", "DocUnits", "ENGLISH", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set DocUnits=ENGLISH, code: {nRet}");
                    return nRet;
                }

                nRet = MTMICRSetValue(strOptions, "ProcessOptions", "DocFeedTimeout", "10000", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set DocFeedTimeout=10000, code: {nRet}");
                    return nRet;
                }

                nRet = MTMICRSetValue(strOptions, "ProcessOptions", "DocFeed", "MANUAL", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set DocFeed=MANUAL, code: {nRet}");
                    return nRet;
                }

                nRet = MTMICRSetValue(strOptions, "ImageOptions", "Number", "2", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set ImageOptions Number=2, code: {nRet}");
                    return nRet;
                }

                nRet = MTMICRSetIndexValue(strOptions, "ImageOptions", "ImageSide", 1, "FRONT", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set ImageSide1=FRONT, code: {nRet}");
                    return nRet;
                }

                nRet = MTMICRSetIndexValue(strOptions, "ImageOptions", "ImageSide", 2, "BACK", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set ImageSide2=BACK, code: {nRet}");
                    return nRet;
                }

                nRet = MTMICRSetIndexValue(strOptions, "ImageOptions", "ImageColor", 1, "COL24", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set ImageColor1=COL24, code: {nRet}");
                    return nRet;
                }

                nRet = MTMICRSetIndexValue(strOptions, "ImageOptions", "ImageColor", 2, "COL24", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set ImageColor2=COL24, code: {nRet}");
                    return nRet;
                }

                nRet = MTMICRSetIndexValue(strOptions, "ImageOptions", "Resolution", 1, "100x100", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set Resolution1=100x100, code: {nRet}");
                    return nRet;
                }

                nRet = MTMICRSetIndexValue(strOptions, "ImageOptions", "Resolution", 2, "100x100", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set Resolution2=100x100, code: {nRet}");
                    return nRet;
                }

                nRet = MTMICRSetIndexValue(strOptions, "ImageOptions", "Compression", 1, "JPEG", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set Compression1=JPEG, code: {nRet}");
                    return nRet;
                }

                nRet = MTMICRSetIndexValue(strOptions, "ImageOptions", "Compression", 2, "JPEG", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set Compression2=JPEG, code: {nRet}");
                    return nRet;
                }

                nRet = MTMICRSetIndexValue(strOptions, "ImageOptions", "FileType", 1, "JPG", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set FileType1=JPG, code: {nRet}");
                    return nRet;
                }

                nRet = MTMICRSetIndexValue(strOptions, "ImageOptions", "FileType", 2, "JPG", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set FileType2=JPG, code: {nRet}");
                    return nRet;
                }

                nRet = MTMICRSetValue(strOptions, "ProcessOptions", "ReadMICR", "E13B", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set ReadMICR=E13B, code: {nRet}");
                    return nRet;
                }

                nRet = MTMICRSetValue(strOptions, "ProcessOptions", "MICRFmt", "6200", ref nActualLength);
                if (nRet != MICR_ST_OK)
                {
                    LogMessage($"SetupOptions: Failed to set MICRFmt=6200, code: {nRet}");
                    return nRet;
                }
            }
            else
            {
                LogMessage($"SetupOptions: Invalid document type: {m_nDocType}");
                return MICR_ST_INVALID_FEED_TYPE;
            }

            m_strOptions = strOptions.ToString();
            LogMessage($"SetupOptions: Options set for {m_nDocType}: {m_strOptions}");
            return MICR_ST_OK;
        }

        private VoucherData ExtractVoucherData()
        {
            StringBuilder strResponse = new StringBuilder(DEFAULT_STRING_BUFFER_SIZE);
            int nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string frontImage = "";
            string frontImagePath = "";
            string backImage = "";
            string backImagePath = "";
            string signature = "";
            string checkDate = "";
            string amount = "";
            string amountWords = "";
            string accountHolder = "";
            string micr = "";
            string checkNumber = "";
            string accountNumber = "";
            string routingNumber = "";
            string bankCode = "";

            if (m_nDocType == DocType.MSR)
            {
                string trackData1 = "";
                string trackData2 = "";
                string trackData3 = "";
                string mpData = "";
                string cardType = "";
                string magnePrintStatus = "";
                string track1Status = "";
                string track2Status = "";
                string track3Status = "";
                string getScore = "";
                string deviceSerialNumber = "";
                string dukptSerialNumber = "";
                string encryptedSessionId = "";
                string encryptedTrack1 = "";
                string encryptedTrack2 = "";
                string encryptedTrack3 = "";

                int nRet = MTMICRGetValue(m_strDocInfo, "MSRInfo", "TrackData1", strResponse, ref nResponseLength);
                trackData1 = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: TrackData1 retrieved: {trackData1}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                nRet = MTMICRGetValue(m_strDocInfo, "MSRInfo", "TrackData2", strResponse, ref nResponseLength);
                trackData2 = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: TrackData2 retrieved: {trackData2}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                nRet = MTMICRGetValue(m_strDocInfo, "MSRInfo", "TrackData3", strResponse, ref nResponseLength);
                trackData3 = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: TrackData3 retrieved: {trackData3}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                nRet = MTMICRGetValue(m_strDocInfo, "MSRInfo", "MPData", strResponse, ref nResponseLength);
                mpData = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: MPData retrieved: {mpData}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                nRet = MTMICRGetValue(m_strDocInfo, "MSRInfo", "CardType", strResponse, ref nResponseLength);
                cardType = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: CardType retrieved: {cardType}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                nRet = MTMICRGetValue(m_strDocInfo, "MSRInfo", "MagnePrintStatus", strResponse, ref nResponseLength);
                magnePrintStatus = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: MagnePrintStatus retrieved: {magnePrintStatus}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                nRet = MTMICRGetValue(m_strDocInfo, "MSRInfo", "Track1Status", strResponse, ref nResponseLength);
                track1Status = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: Track1Status retrieved: {track1Status}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                nRet = MTMICRGetValue(m_strDocInfo, "MSRInfo", "Track2Status", strResponse, ref nResponseLength);
                track2Status = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: Track2Status retrieved: {track2Status}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                nRet = MTMICRGetValue(m_strDocInfo, "MSRInfo", "Track3Status", strResponse, ref nResponseLength);
                track3Status = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: Track3Status retrieved: {track3Status}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                nRet = MTMICRGetValue(m_strDocInfo, "MSRInfo", "GetScore", strResponse, ref nResponseLength);
                getScore = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: GetScore retrieved: {getScore}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                nRet = MTMICRGetValue(m_strDocInfo, "DeviceInfo", "DeviceSerialNumber", strResponse, ref nResponseLength);
                deviceSerialNumber = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: DeviceSerialNumber retrieved: {deviceSerialNumber}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                nRet = MTMICRGetValue(m_strDocInfo, "MSRInfo", "DUKPTSerialNumber", strResponse, ref nResponseLength);
                dukptSerialNumber = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: DUKPTSerialNumber retrieved: {dukptSerialNumber}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                nRet = MTMICRGetValue(m_strDocInfo, "MSRInfo", "EncryptedSessionID", strResponse, ref nResponseLength);
                encryptedSessionId = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: EncryptedSessionID retrieved: {encryptedSessionId}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                nRet = MTMICRGetValue(m_strDocInfo, "MSRInfo", "EncryptedTrack1", strResponse, ref nResponseLength);
                encryptedTrack1 = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: EncryptedTrack1 retrieved: {encryptedTrack1}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                nRet = MTMICRGetValue(m_strDocInfo, "MSRInfo", "EncryptedTrack2", strResponse, ref nResponseLength);
                encryptedTrack2 = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: EncryptedTrack2 retrieved: {encryptedTrack2}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                nRet = MTMICRGetValue(m_strDocInfo, "MSRInfo", "EncryptedTrack3", strResponse, ref nResponseLength);
                encryptedTrack3 = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: EncryptedTrack3 retrieved: {encryptedTrack3}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                return new VoucherData
                {
                    voucherNo = trackData2,
                    voucherType = cardType,
                    micr = "",
                    frontImage = "",
                    backImage = "",
                    narration = "", // Narration is set by frontend
                    frontImagePath = "",
                    backImagePath = "",
                    trackData1 = trackData1,
                    trackData2 = trackData2,
                    trackData3 = trackData3,
                    mpData = mpData,
                    cardType = cardType,
                    magnePrintStatus = magnePrintStatus,
                    track1Status = track1Status,
                    track2Status = track2Status,
                    track3Status = track3Status,
                    getScore = getScore,
                    deviceSerialNumber = deviceSerialNumber,
                    dukptSerialNumber = dukptSerialNumber,
                    encryptedSessionId = encryptedSessionId,
                    encryptedTrack1 = encryptedTrack1,
                    encryptedTrack2 = encryptedTrack2,
                    encryptedTrack3 = encryptedTrack3,
                    checkNumber = "",
                    accountNumber = "",
                    routingNumber = "",
                    bankCode = "",
                    checkDate = "",
                    amount = "",
                    amountWords = "",
                    accountHolder = "",
                    signature = ""
                };
            }
            else
            {
                // VoucherNo is passed via URL, not derived from MICR
                string voucherNo = Request.Path.Value?.Split('/').Last() ?? "";

                int nRet = MTMICRGetValue(m_strDocInfo, "DocInfo", "MICRRaw", strResponse, ref nResponseLength);
                micr = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: MICRRaw retrieved: {micr}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                // Parse MICR manually
                if (!string.IsNullOrEmpty(micr))
                {
                    try
                    {
                        var parts = micr.Split(new[] { 'U', 'T' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length >= 4)
                        {
                            checkNumber = parts[0].Trim(); // Check No. (e.g., 000002)
                            routingNumber = parts[1].Replace("?", "").Trim(); // Routing No. (e.g., 90109)
                            accountNumber = parts[2].Trim(); // Account No. (e.g., 9040007857211)
                            bankCode = parts[3].Trim(); // Bank Code (e.g., 01)
                        }
                        LogMessage($"ExtractVoucherData: Parsed MICR - CheckNo: {checkNumber}, RoutingNo: {routingNumber}, AccountNo: {accountNumber}, BankCode: {bankCode}");
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"ExtractVoucherData: MICR parsing error: {ex.Message}");
                    }
                }

                nRet = MTMICRGetValue(m_strDocInfo, "DocInfo", "AccountNumber", strResponse, ref nResponseLength);
                if (string.IsNullOrEmpty(accountNumber)) accountNumber = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: AccountNumber retrieved: {accountNumber}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                nRet = MTMICRGetValue(m_strDocInfo, "DocInfo", "RoutingNumber", strResponse, ref nResponseLength);
                if (string.IsNullOrEmpty(routingNumber)) routingNumber = strResponse.ToString().Trim();
                LogMessage($"ExtractVoucherData: RoutingNumber retrieved: {routingNumber}, RetCode={nRet}");
                nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                strResponse.Clear();

                byte[]? frontImageBuf = null;
                const int maxRetries = 5;
                for (int attempt = 1; attempt <= maxRetries; attempt++)
                {
                    int nFrontImageSize = 0;
                    string strFrontImageID = "";
                    nRet = MTMICRGetIndexValue(m_strDocInfo, "ImageInfo", "ImageSize", 1, strResponse, ref nResponseLength);
                    if (nRet != MICR_ST_OK)
                    {
                        LogMessage($"ExtractVoucherData: Failed to get FrontImageSize on attempt {attempt}, RetCode={nRet}");
                        nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                        strResponse.Clear();
                        continue;
                    }
                    nFrontImageSize = int.TryParse(strResponse.ToString(), out int size) ? size : 0;
                    LogMessage($"ExtractVoucherData: FrontImageSize retrieved: {nFrontImageSize}, RetCode={nRet}");
                    nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                    strResponse.Clear();

                    if (nFrontImageSize > 0)
                    {
                        nRet = MTMICRGetIndexValue(m_strDocInfo, "ImageInfo", "ImageURL", 1, strResponse, ref nResponseLength);
                        if (nRet != MICR_ST_OK)
                        {
                            LogMessage($"ExtractVoucherData: Failed to get FrontImageURL on attempt {attempt}, RetCode={nRet}");
                            nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                            strResponse.Clear();
                            continue;
                        }
                        strFrontImageID = strResponse.ToString().Trim();
                        LogMessage($"ExtractVoucherData: FrontImageURL retrieved: {strFrontImageID}, RetCode={nRet}");
                        nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                        strResponse.Clear();

                        if (!string.IsNullOrEmpty(strFrontImageID))
                        {
                            frontImageBuf = new byte[nFrontImageSize];
                            int nImageLength = nFrontImageSize;
                            nRet = MTMICRGetImage(m_strCurrentDeviceName, strFrontImageID, frontImageBuf, ref nImageLength);
                            LogMessage($"ExtractVoucherData: FrontImage attempt {attempt} MTMICRGetImage returned {nRet}, Size={nImageLength}");
                            if (nRet == MICR_ST_OK && nImageLength > 0)
                            {
                                frontImage = Convert.ToBase64String(frontImageBuf, 0, nImageLength);
                                LogMessage($"ExtractVoucherData: FrontImage Base64 length={frontImage.Length}");
                                try
                                {
                                    frontImagePath = Path.Combine(imageSavePath, $"front_{timestamp}.jpg");
                                    LogMessage($"ExtractVoucherData: Attempting to save front image to {frontImagePath}");
                                    System.IO.File.WriteAllBytes(frontImagePath, frontImageBuf.Take(nImageLength).ToArray());
                                    LogMessage($"ExtractVoucherData: Saved front image to {frontImagePath}, Size={nImageLength} bytes");
                                }
                                catch (Exception ex)
                                {
                                    LogMessage($"ExtractVoucherData: Error saving front image to {frontImagePath}: {ex.Message}, StackTrace: {ex.StackTrace}");
                                }
                                break;
                            }
                            else
                            {
                                LogMessage($"ExtractVoucherData: Failed to get FrontImage on attempt {attempt}, RetCode={nRet}, ImageLength={nImageLength}");
                            }
                        }
                    }
                    if (attempt < maxRetries)
                    {
                        Thread.Sleep(1000);
                    }
                }

                // Save raw front image if available
                if (frontImageBuf != null)
                {
                    try
                    {
                        frontImagePath = Path.Combine(imageSavePath, $"raw_front_{timestamp}.jpg");
                        System.IO.File.WriteAllBytes(frontImagePath, frontImageBuf);
                        LogMessage($"ExtractVoucherData: Saved raw front image to {frontImagePath}");
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"ExtractVoucherData: Error saving raw front image: {ex.Message}, StackTrace: {ex.StackTrace}");
                    }
                }

                for (int attempt = 1; attempt <= maxRetries; attempt++)
                {
                    int nBackImageSize = 0;
                    string strBackImageID = "";
                    nRet = MTMICRGetIndexValue(m_strDocInfo, "ImageInfo", "ImageSize", 2, strResponse, ref nResponseLength);
                    if (nRet != MICR_ST_OK)
                    {
                        LogMessage($"ExtractVoucherData: Failed to get BackImageSize on attempt {attempt}, RetCode={nRet}");
                        nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                        strResponse.Clear();
                        continue;
                    }
                    nBackImageSize = int.TryParse(strResponse.ToString(), out int backSize) ? backSize : 0;
                    LogMessage($"ExtractVoucherData: BackImageSize retrieved: {nBackImageSize}, RetCode={nRet}");
                    nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                    strResponse.Clear();

                    if (nBackImageSize > 0)
                    {
                        nRet = MTMICRGetIndexValue(m_strDocInfo, "ImageInfo", "ImageURL", 2, strResponse, ref nResponseLength);
                        if (nRet != MICR_ST_OK)
                        {
                            LogMessage($"ExtractVoucherData: Failed to get BackImageURL on attempt {attempt}, RetCode={nRet}");
                            nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                            strResponse.Clear();
                            continue;
                        }
                        strBackImageID = strResponse.ToString().Trim();
                        LogMessage($"ExtractVoucherData: BackImageURL retrieved: {strBackImageID}, RetCode={nRet}");
                        nResponseLength = DEFAULT_STRING_BUFFER_SIZE;
                        strResponse.Clear();

                        if (!string.IsNullOrEmpty(strBackImageID))
                        {
                            byte[] backImageBuf = new byte[nBackImageSize];
                            int nImageLength = nBackImageSize;
                            nRet = MTMICRGetImage(m_strCurrentDeviceName, strBackImageID, backImageBuf, ref nImageLength);
                            LogMessage($"ExtractVoucherData: BackImage attempt {attempt} MTMICRGetImage returned {nRet}, Size={nImageLength}");
                            if (nRet == MICR_ST_OK && nImageLength > 0)
                            {
                                backImage = Convert.ToBase64String(backImageBuf, 0, nImageLength);
                                LogMessage($"ExtractVoucherData: BackImage Base64 length={backImage.Length}");
                                try
                                {
                                    backImagePath = Path.Combine(imageSavePath, $"back_{timestamp}.jpg");
                                    LogMessage($"ExtractVoucherData: Attempting to save back image to {backImagePath}");
                                    System.IO.File.WriteAllBytes(backImagePath, backImageBuf.Take(nImageLength).ToArray());
                                    LogMessage($"ExtractVoucherData: Saved back image to {backImagePath}, Size={nImageLength} bytes");
                                }
                                catch (Exception ex)
                                {
                                    LogMessage($"ExtractVoucherData: Error saving back image to {backImagePath}: {ex.Message}, StackTrace: {ex.StackTrace}");
                                }
                                break;
                            }
                            else
                            {
                                LogMessage($"ExtractVoucherData: Failed to get BackImage on attempt {attempt}, RetCode={nRet}, ImageLength={nImageLength}");
                            }
                        }
                    }
                    if (attempt < maxRetries)
                    {
                        Thread.Sleep(1000);
                    }
                }

                return new VoucherData
                {
                    voucherNo = voucherNo,
                    voucherType = "",
                    micr = micr,
                    frontImage = frontImage,
                    backImage = backImage,
                    narration = "", // Narration is set by frontend
                    frontImagePath = frontImagePath,
                    backImagePath = backImagePath,
                    trackData1 = "",
                    trackData2 = "",
                    trackData3 = "",
                    mpData = "",
                    cardType = "",
                    magnePrintStatus = "",
                    track1Status = "",
                    track2Status = "",
                    track3Status = "",
                    getScore = "",
                    deviceSerialNumber = "",
                    dukptSerialNumber = "",
                    encryptedSessionId = "",
                    encryptedTrack1 = "",
                    encryptedTrack2 = "",
                    encryptedTrack3 = "",
                    checkNumber = checkNumber,
                    accountNumber = accountNumber,
                    routingNumber = routingNumber,
                    bankCode = bankCode,
                    checkDate = checkDate,
                    amount = amount,
                    amountWords = amountWords,
                    accountHolder = accountHolder,
                    signature = signature
                };
            }
        }

        public class VoucherData
        {
            public string voucherNo { get; set; } = "";
            public string voucherType { get; set; } = "";
            public string micr { get; set; } = "";
            public string frontImage { get; set; } = "";
            public string backImage { get; set; } = "";
            public string narration { get; set; } = "";
            public string frontImagePath { get; set; } = "";
            public string backImagePath { get; set; } = "";
            public string trackData1 { get; set; } = "";
            public string trackData2 { get; set; } = "";
            public string trackData3 { get; set; } = "";
            public string mpData { get; set; } = "";
            public string cardType { get; set; } = "";
            public string magnePrintStatus { get; set; } = "";
            public string track1Status { get; set; } = "";
            public string track2Status { get; set; } = "";
            public string track3Status { get; set; } = "";
            public string getScore { get; set; } = "";
            public string deviceSerialNumber { get; set; } = "";
            public string dukptSerialNumber { get; set; } = "";
            public string encryptedSessionId { get; set; } = "";
            public string encryptedTrack1 { get; set; } = "";
            public string encryptedTrack2 { get; set; } = "";
            public string encryptedTrack3 { get; set; } = "";
            public string checkNumber { get; set; } = ""; // From MICR, not voucherNo
            public string accountNumber { get; set; } = "";
            public string routingNumber { get; set; } = "";
            public string bankCode { get; set; } = "";
            public string checkDate { get; set; } = "";
            public string amount { get; set; } = ""; // Figures
            public string amountWords { get; set; } = ""; // Words from left section
            public string accountHolder { get; set; } = "";
            public string signature { get; set; } = "";
        }

        #region DLL Imports
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern int CreateFile(string lpFileName, uint dwDesiredAccess, uint dwShareMode,
            uint lpSecurityAttributes, uint dwCreationDisposition, uint dwFlagsAndAttributes, int hTemplateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(int hHandle);

        [DllImport("mtxmlmcr.dll", SetLastError = true)]
        private static extern int MTMICRGetImage(string strDeviceName, string strImageID, byte[] imageBuf, ref int nBufLength);

        [DllImport("mtxmlmcr.dll")]
        private static extern int MTMICRGetDevice(int nDeviceIndex, StringBuilder strDeviceName);

        [DllImport("mtxmlmcr.dll")]
        private static extern int MTMICRQueryInfo(string strDeviceName, string strQueryParm, StringBuilder strResponse, ref int nResponseLength);

        [DllImport("mtxmlmcr.dll")]
        private static extern int MTMICRSetValue(StringBuilder strOptions, string strSection, string strKey, string strValue, ref int nActualLength);

        [DllImport("mtxmlmcr.dll")]
        private static extern int MTMICRGetValue(string strDocInfo, string strSection, string strKey, StringBuilder strResponse, ref int nResponseLength);

        [DllImport("mtxmlmcr.dll")]
        private static extern int MTMICRSetIndexValue(StringBuilder strOptions, string strSection, string strKey, int nIndex, string strValue, ref int nActualLength);

        [DllImport("mtxmlmcr.dll")]
        private static extern int MTMICRGetIndexValue(string strDocInfo, string strSection, string strKey, int nIndex, StringBuilder strResponse, ref int nResponseLength);

        [DllImport("mtxmlmcr.dll")]
        private static extern int MTMICRProcessCheck(string strDeviceName, string strOptions, StringBuilder strResponse, ref int nResponseLength);

        [DllImport("mtxmlmcr.dll")]
        private static extern int MTMICRSetLogFileHandle(int hLogFile);

        [DllImport("mtxmlmcr.dll")]
        private static extern int MTMICROpenDevice(string strDeviceName);

        [DllImport("mtxmlmcr.dll")]
        private static extern int MTMICRCloseDevice(string strDeviceName);
        #endregion

        #region Constants
        private const uint GENERIC_READ = 0x80000000;
        private const uint GENERIC_WRITE = 0x40000000;
        private const uint FILE_SHARE_READ = 0x00000001;
        private const uint FILE_SHARE_WRITE = 0x00000002;
        private const uint OPEN_ALWAYS = 4;
        private const uint FILE_ATTRIBUTE_NORMAL = 0x00000080;
        private const int MICR_ST_OK = 0;
        private const int MICR_ST_DEVICE_NOT_FOUND = -7;
        private const int MICR_ST_INVALID_FEED_TYPE = -17;
        #endregion
    }
}