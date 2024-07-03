//
// MovingAverage.cs
//
// Author: Jeffrey Stedfast <jestedfa@microsoft.com>
//
// Copyright (c) 2016-2024 Jeffrey Stedfast
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using Newtonsoft.Json;

namespace IvyPortfolio
{
	public enum MovingAverageAlgorithm
	{
		Simple = 1
	}

	public enum MovingAveragePeriodType
	{
		Day = 1,
		//Week = 7,
		Month = 28,
		//Year = 365
	}

	public class MovingAverage
	{
		[JsonProperty ("algorithm")]
		public MovingAverageAlgorithm Algorithm { get; set; }

		[JsonProperty ("period-type")]
		public MovingAveragePeriodType PeriodType { get; set; }

		[JsonProperty ("period")]
		public int Period { get; set; }

		[JsonProperty ("title")]
		public string Title { get; set; }
	}

	public class MovingAverageConverter : JsonConverter
	{
		readonly Dictionary<string, MovingAverageAlgorithm> algorithms;
		readonly Dictionary<string, MovingAveragePeriodType> periodTypes;

		public MovingAverageConverter ()
		{
			algorithms = new Dictionary<string, MovingAverageAlgorithm> (StringComparer.OrdinalIgnoreCase);
			foreach (MovingAverageAlgorithm algorithm in Enum.GetValues (typeof (MovingAverageAlgorithm)))
				algorithms.Add (algorithm.ToString (), algorithm);

			periodTypes = new Dictionary<string, MovingAveragePeriodType> (StringComparer.OrdinalIgnoreCase);
			foreach (MovingAveragePeriodType type in Enum.GetValues (typeof (MovingAveragePeriodType)))
				periodTypes.Add (type.ToString (), type);
		}

		public override bool CanConvert (Type objectType)
		{
			return objectType == typeof (MovingAverageAlgorithm) || objectType == typeof (MovingAveragePeriodType);
		}

		public override object ReadJson (JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
		{
			var value = reader.Value?.ToString ();

			if (objectType == typeof (MovingAverageAlgorithm)) {
				if (value == null || !algorithms.TryGetValue (value, out var algorithm))
					return MovingAverageAlgorithm.Simple;

				return algorithm;
			}

			if (objectType == typeof (MovingAveragePeriodType)) {
				if (value == null || !periodTypes.TryGetValue (value, out var periodType))
					return MovingAveragePeriodType.Day;

				return periodType;
			}

			throw new NotImplementedException ();
		}

		public override void WriteJson (JsonWriter writer, object value, JsonSerializer serializer)
		{
			throw new NotImplementedException ();
		}
	}
}
