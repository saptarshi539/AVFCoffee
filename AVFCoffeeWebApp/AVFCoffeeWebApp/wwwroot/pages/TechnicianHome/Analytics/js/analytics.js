var lang = localStorage.getItem("selectedLanguage");
//$(document).on('ready', function () {
//    console.log(lang);
//    GetAnalyses(apiURL);
//    GetFarms(apiURL);
//});

//function GetAnalyses(apiURL) {
//    $.ajax({
//        type: "GET",
//        url: apiURL + "TechnicianHomeAPI/GetAnalysis",
//        //data: "language=" + language,
//        contentType: "application/json",
//        success: function (result) {
//            console.log(result);
//        }
//    });
//}

//function GetFarms(apiURL) {
//    $.ajax({
//        type: "GET",
//        url: apiURL + "TechnicianHomeAPI/GetFarms",
//        //data: "language=" + language,
//        contentType: "application/json",
//        success: function (result) {
//            console.log(result);
//        }
//    });
//}

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var Card = function (_React$Component) {
    _inherits(Card, _React$Component);

    function Card(props) {
        _classCallCheck(this, Card);

        var _this = _possibleConstructorReturn(this, (Card.__proto__ || Object.getPrototypeOf(Card)).call(this, props));

        _this.state = {};
        return _this;
    }

    _createClass(Card, [{
        key: "renderBadge",
        value: function renderBadge(input_length) {
            if (input_length > 2) {
                return React.createElement(
                    "div",
                    { className: "more-badge" },
                    "+",
                    input_length - 2
                );
            }
        }
    }, {
        key: "render",
        value: function render() {

            return React.createElement(
                "div",
                { className: "doc-card" },
                React.createElement("img", { className: "doc-card-more-icon", src: "more-vert.svg", width: "24", height: "24", alt: "more options" }),
                React.createElement(
                    "div",
                    { className: "doc-card-title" },
                    this.props.info.title
                ),
                React.createElement(
                    "div",
                    { className: "doc-card-subtitles" },
                    React.createElement(
                        "div",
                        { className: "doc-card-sub-title" },
                        this.props.info.input[0]
                    ),
                    React.createElement(
                        "div",
                        { className: "doc-card-sub-title" },
                        this.props.info.input[1]
                    ),
                    this.renderBadge(this.props.info.input.length)
                ),
                React.createElement(
                    "div",
                    { className: "doc-card-meta" },
                    "last modified: yesterday"
                )
            );
        }
    }]);

    return Card;
}(React.Component);

var DocView = function (_React$Component2) {
    _inherits(DocView, _React$Component2);

    function DocView(props) {
        _classCallCheck(this, DocView);

        var _this2 = _possibleConstructorReturn(this, (DocView.__proto__ || Object.getPrototypeOf(DocView)).call(this, props));

        _this2.state = {
            error: null,
            isLoaded: true,
            analyses: []
        };
        return _this2;
    }

    _createClass(DocView, [{
        key: "componentDidMount",
        value: function componentDidMount() {
            var _this3 = this;

            fetch(apiURL + "TechnicianHomeAPI/GetAnalysis").then(function (res) {
                return res.json();
            }).then(function (result) {
                _this3.setState({
                    isLoaded: true,
                    analyses: result.analyses
                });
            },
                // Note: it's important to handle errors here
                // instead of a catch() block so that we don't swallow
                // exceptions from actual bugs in components.
                function (error) {
                    _this3.setState({
                        isLoaded: true,
                        error: error
                    });
                });
        }
    }, {
        key: "renderCards",
        value: function renderCards(data) {
            var cards = data.map(function (analysis, index) {
                return React.createElement(Card, { key: index, info: analysis });
            });

            return React.createElement(
                "ul",
                { className: "doc-card-list" },
                cards
            );
        }
    }, {
        key: "render",
        value: function render() {
            // TODO: Get data for all analyses (for each card)
            var data = {
                analyses: [{
                    title: 'Verde District Comparison',
                    input: ['GARCIA CARRASCO LUIS ALBERTO', 'VERDE DISTRICT', 'Shouldnt show'],
                    timeStamp: '8/31/2018 7:08:30 PM'
                }, {
                    title: '1:1 Bermeo',
                    input: ['BERMEO GARCIA PABLO', 'HIGH ALTITUDE FARMS'],
                    timeStamp: '8/31/2018 7:08:30 PM'
                }, {
                    title: 'Organic Farm Performace',
                    input: ['ORGANIC FARMS'],
                    timeStamp: '8/31/2018 7:08:30 PM'
                }]
            };
            var _state = this.state,
                error = _state.error,
                isLoaded = _state.isLoaded,
                analyses = _state.analyses;

            if (error) {
                return React.createElement(
                    "div",
                    null,
                    "Error: ",
                    error.message
                );
            } else if (!isLoaded) {
                return React.createElement(
                    "div",
                    null,
                    "Loading..."
                );
            } else {
                return React.createElement(
                    "div",
                    { id: "doc-cards-view" },
                    this.renderCards(this.state.analyses),
                    React.createElement(
                        "button",
                        { className: "add-FAB", type: "button" },
                        React.createElement("img", { src: "./icons/add-icon.svg", alt: "add analysis" })
                    )
                );
            }
        }
    }]);

    return DocView;
}(React.Component);

var Analytics = function (_React$Component3) {
    _inherits(Analytics, _React$Component3);

    function Analytics(props) {
        _classCallCheck(this, Analytics);

        var _this4 = _possibleConstructorReturn(this, (Analytics.__proto__ || Object.getPrototypeOf(Analytics)).call(this, props));

        _this4.state = {};
        return _this4;
    }

    _createClass(Analytics, [{
        key: "render",
        value: function render() {
            return React.createElement(DocView, null);
        }
    }]);

    return Analytics;
}(React.Component);

ReactDOM.render(React.createElement(Analytics, null), document.getElementById("analytics-main"));;