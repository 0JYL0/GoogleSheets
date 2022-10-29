using Google.Apis.Sheets.v4.Data;
using Google.Apis.Sheets.v4;
using Microsoft.AspNetCore.Mvc;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;
using System.Linq;
using System.Collections.Generic;
using System;

namespace GoogleSheets_1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ItemsController : ControllerBase
    {
        const string SPREADSHEET_ID = "1Tg_GUhaJBzAY8s-wjwV9zidxjj0DCXrcmgLzaxfHt2o";
        const string SHEET_NAME = "Item";

        SpreadsheetsResource.ValuesResource _googleSheetValues;

        public ItemsController(GoogleSheetsHelper googleSheetsHelper)
        {
            _googleSheetValues = googleSheetsHelper.Service.Spreadsheets.Values;
        }

        [HttpGet]
        public IActionResult Get()
        {
            var range = $"A:D";

            var request = _googleSheetValues.Get(SPREADSHEET_ID, range);
            var response = request.Execute();
            var values = response.Values;

            return Ok(ItemsMapper.MapFromRangeData(values));
        }

        [HttpGet("{rowId}")]
        public IActionResult Get(int rowId)
        {
            var range = $"A{rowId}:D{rowId}";
            var request = _googleSheetValues.Get(SPREADSHEET_ID, range);
            var response = request.Execute();
            var values = response.Values;

            return Ok(ItemsMapper.MapFromRangeData(values).FirstOrDefault());
        }

        [HttpPost]
        public IActionResult Post(List<Item> item)
        {
            try
            {

                var range = $"A:D";
                var valueRange = new ValueRange
                {
                    Values = ItemsMapper.MapToRangeData(item)
                };

                var appendRequest = _googleSheetValues.Append(valueRange, SPREADSHEET_ID, range);
                appendRequest.ValueInputOption = AppendRequest.ValueInputOptionEnum.RAW;
                appendRequest.InsertDataOption = AppendRequest.InsertDataOptionEnum.INSERTROWS;
                appendRequest.Execute();
            }
            catch(Exception e)
            {
                return BadRequest(e.Message);
            }

            return CreatedAtAction(nameof(Get), item);
        }

        [HttpPut("{rowId}")]
        public IActionResult Put(int rowId, List<Item> item)
        {
            var range = $"A{rowId}:D{rowId}";
            var valueRange = new ValueRange
            {
                Values = ItemsMapper.MapToRangeData(item)
            };

            var updateRequest = _googleSheetValues.Update(valueRange, SPREADSHEET_ID, range);
            updateRequest.ValueInputOption = UpdateRequest.ValueInputOptionEnum.USERENTERED;
            updateRequest.Execute();

            return NoContent();
        }

        [HttpDelete("{rowId}")]
        public IActionResult Delete(int rowId)
        {
            var range = $"A{rowId}:D{rowId}";
            var requestBody = new ClearValuesRequest();

            var deleteRequest = _googleSheetValues.Clear(requestBody, SPREADSHEET_ID, range);
            deleteRequest.Execute();

            return NoContent();
        }

        [HttpDelete("")]
        public IActionResult DeleteAll()
        {
            var request = _googleSheetValues.Get(SPREADSHEET_ID, "A:D");
            var response = request.Execute();
            var val = response.Values.Count;

            var range = $"A{2}:D{val}";
            var requestBody = new ClearValuesRequest();

            var deleteRequest = _googleSheetValues.Clear(requestBody, SPREADSHEET_ID, range);
            deleteRequest.Execute();

            return NoContent();
        }
    }
}
