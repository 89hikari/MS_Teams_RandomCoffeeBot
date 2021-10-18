// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
    TurnContext,
    MessageFactory,
    TeamsInfo,
    TeamsActivityHandler,
    CardFactory,
    ActionTypes,
} = require('botbuilder');

const sleep = require('util').promisify(setTimeout)

const ACData = require("adaptivecards-templating");

const rawWelcomeCard = require("./adaptiveCards/welcome.json");
const yesOrNoCard = require("./adaptiveCards/yesOrNoCard.json");
const checkConversationCard = require("./adaptiveCards/checkConversation.json");
const gradeCard = require("./adaptiveCards/gradeCard.json")

const TextEncoder = require('util').TextEncoder;

class TeamsConversationBot extends TeamsActivityHandler {
    constructor() {
        super();

        this.data_object = {
            startDate: "",
            endDate: "",
            comment: ""
        }

        this.team_info = {
            id: "",
            name: "",
            aadGroupId: "",
            channelCount: 0,
            memberCount: 0,
            tenantId: ""
        }

        this.if_started = false;

        this.groupContext = {};

        this.notConversed = [];

        this.personList = [];

        this.denied_persons = [];

        this.contextes = [];

        this.pairs = [];

        this.conversedList = [];

        this.notConversed = [];

        this.participatesCount = 0;

        this.talkedCount = 0;

        this.tmpArr = [{
            id: '29:1KLUufK2i58IsYP0_Z4hc3UdratvN7zcqV8buLvsu3GoDW87u6w7brr9a3cE6DvgNL7RFeHXG874jh2jPGGPCSA',
            name: 'Белоусов Владислав',
            aadObjectId: '694b9c26-a8a9-4828-8746-8e7c4544a943'
        }, {
            id: '29:1KLUufK2i58IsYP0_Z4hc3UdratvN7zcqV8buLvsu3GoDW87u6w7brr9a3cE6DvgNL7RFeHXG874jh2jPGGPCSA',
            name: 'Белоусов Владислав',
            aadObjectId: '694b9c26-a8a9-4828-8746-8e7c4544a943'
        },
        {
            id: '29:1KLUufK2i58IsYP0_Z4hc3UdratvN7zcqV8buLvsu3GoDW87u6w7brr9a3cE6DvgNL7RFeHXG874jh2jPGGPCSA',
            name: 'Белоусов Владислав',
            aadObjectId: '694b9c26-a8a9-4828-8746-8e7c4544a943'
        }, {
            id: '29:1KLUufK2i58IsYP0_Z4hc3UdratvN7zcqV8buLvsu3GoDW87u6w7brr9a3cE6DvgNL7RFeHXG874jh2jPGGPCSA',
            name: 'Белоусов Владислав',
            aadObjectId: '694b9c26-a8a9-4828-8746-8e7c4544a943'
        }];

        this.gradesAndComments = {
            grades: [],
            comments: []
        }

        function shuffle(array) {
            let currentIndex = array.length, randomIndex;

            while (currentIndex != 0) {

                randomIndex = Math.floor(Math.random() * currentIndex);
                currentIndex--;

                [array[currentIndex], array[randomIndex]] = [
                    array[randomIndex], array[currentIndex]];
            }

            return array;
        }

        this.onMessage(async (context, next) => {
            TurnContext.removeRecipientMention(context.activity);

            if (context.activity.text != undefined) {
                try {

                    var if_reset = context.activity.text.trim().toLocaleLowerCase();

                    if (if_reset.includes('emergency')) {
                        await this.emergencyReset();
                    }

                } catch (error) {
                    console.log("Паника на борту!!1!!!!11111!");
                }
            }

            if (context.activity.value != undefined) {

                if (context.activity.value.grade != undefined && context.activity.value.grade != "" && context.activity.value.grade_text != "" && context.activity.value.grade_text != undefined) {
                    this.gradesAndComments.grades.push(context.activity.value.grade);
                    this.gradesAndComments.comments.push(context.activity.value.grade_text);
                    await this.thanksForParticipating(context);
                    await this.deleteCardActivityAsync(context);
                }

                if (context.activity.value.startDate != undefined && context.activity.value.startDate != "" && context.activity.value.endDate != undefined &&
                    context.activity.value.endDate != "" && context.activity.value.comment != "" && context.activity.value.comment != undefined) {

                    this.groupContext = context;

                    var dayStart = context.activity.value.startDate.slice(8, context.activity.value.startDate.length);
                    var monthStart = context.activity.value.startDate.slice(5, 7);
                    var yearStart = context.activity.value.startDate.slice(2, 4);

                    var dayEnd = context.activity.value.endDate.slice(8, context.activity.value.startDate.length);
                    var monthEnd = context.activity.value.endDate.slice(5, 7);
                    var yearEnd = context.activity.value.endDate.slice(2, 4);

                    this.data_object.startDate = dayStart + "." + monthStart + "." + yearStart;
                    this.data_object.endDate = dayEnd + "." + monthEnd + "." + yearEnd;
                    this.data_object.comment = context.activity.value.comment;

                    var yearNum = +("20" + yearStart);
                    var monthNum = monthStart[0] == "0" ? +monthStart[1] - 1 : +monthStart - 1;
                    var dayNum = dayStart[0] == "0" ? +dayStart[1] : +dayStart;

                    var yearNumEnd = +("20" + yearEnd);
                    var monthNumEnd = monthEnd[0] == "0" ? +monthEnd[1] - 1 : +monthEnd - 1;
                    var dayNumEnd = dayEnd[0] == "0" ? +dayEnd[1] : +dayEnd;

                    //year, month 0-11, date, hour, min (can add ,sec,msec)
                    var eta_ms = new Date(yearNum, monthNum, dayNum, 10, 0).getTime() - Date.now();

                    var eta_ms_end = new Date(yearNumEnd, monthNumEnd, dayNumEnd, 9, 0).getTime() - Date.now() - 518400000;

                    // console.log((Math.ceil(eta_ms_end / 86400000)) * 86400000);

                    await this.deleteCardActivityAsync(context);

                    try {
                        const message = MessageFactory.text(`Параметры заданы. Дата начала: ${this.data_object.startDate}, дата окончания: ${this.data_object.endDate}. Тема: ${this.data_object.comment}.`);
                        await context.sendActivity(message);
                    } catch (e) {
                        if (e.code === 'MemberNotFoundInConversation') {
                            return context.sendActivity(MessageFactory.text('Member not found.'));
                        } else {
                            throw e;
                        }
                    }

                    await sleep(eta_ms);

                    await this.messageAllMembersAsync(context);

                    await sleep(259200000); // тут await sleep(259200000);

                    //////////

                    shuffle(this.tmpArr);

                    // this.personList = this.tmpArr; // На время

                    //////////

                    var tmp_list = [];
                    if (this.personList.length >= 2) {
                        if (this.personList.length % 2 == 0) {
                            for (let i = 0; i < this.personList.length; i++) {
                                tmp_list.push(this.personList[i]);
                                if (i % 2 != 0 && i != 0) {
                                    this.pairs.push(tmp_list);
                                    tmp_list = [];
                                }
                            }
                        } else {
                            for (let i = 0; i < this.personList.length; i++) {
                                tmp_list.push(this.personList[i]);
                                if (i == this.personList.length - 1) {
                                    tmp_list.push(this.personList[i]);
                                    this.pairs.push(tmp_list);
                                    break;
                                }
                                if (i % 2 != 0 && i != 0) {
                                    this.pairs.push(tmp_list);
                                    tmp_list = [];
                                }
                            }
                        }
                    }

                    this.if_started = true;

                    await this.messageToPair(context);

                    await await sleep((Math.ceil(eta_ms_end / 86400000)) * 86400000); // тут await sleep((Math.ceil(eta_ms_end / 86400000)) * 86400000);

                    await this.checkIfConversationWasAThing(context);

                    await sleep(259200000); // тут await sleep(259200000);

                    await this.messageToDeniedPersons(context);

                    await sleep(259200000); // тут await sleep(259200000);

                    await this.sendInfoAboutLastEvent(context);
                }

                await next();

            } else {

                const text = context.activity.text.trim().toLocaleLowerCase();

                if (text.includes('start')) {
                    if (this.data_object.endDate == "" || this.data_object.startDate == "" || this.data_object.comment == "") {
                        await this.cardActivityAsync(context, false);
                    }
                }

                if (this.data_object.endDate != "" && this.data_object.startDate != "" && this.data_object.comment != "") {

                    if (text == 'в этот раз воздержусь' && this.data_object.startDate != "" && this.data_object.endDate != "" && this.data_object.comment != "" && !this.if_started) {
                        await this.cutItOff(context);
                    } else if (text == 'да, я с вами' && this.data_object.startDate != "" && this.data_object.endDate != "" && this.data_object.comment != "" && !this.if_started) {
                        await this.savePerson(context);
                    } else if (text == 'нет ещё' && this.data_object.startDate != "" && this.data_object.endDate != "" && this.data_object.comment != "" && this.if_started) {
                        await this.notYet(context);
                    } else if (text == 'да, мы поболтали' && this.data_object.startDate != "" && this.data_object.endDate != "" && this.data_object.comment != "" && this.if_started) {
                        await this.gradeAndComment(context);
                    }
                }

                await next();
            }

        });

        this.onMembersAddedActivity(async (context, next) => {
            await Promise.all((context.activity.membersAdded || []).map(async (member) => {
                if (member.id !== context.activity.recipient.id) {
                    await context.sendActivity(
                        `Welcome to the team ${member.givenName} ${member.surname}`
                    );
                }
            }));

            await next();
        });

        this.onReactionsAdded(async (context) => {
            await Promise.all((context.activity.reactionsAdded || []).map(async (reaction) => {
                const newReaction = `You reacted with '${reaction.type}' to the following message: '${context.activity.replyToId}'`;
                const resourceResponse = await context.sendActivity(newReaction);
                // Save information about the sent message and its ID (resourceResponse.id).
            }));
        });

        this.onReactionsRemoved(async (context) => {
            await Promise.all((context.activity.reactionsRemoved || []).map(async (reaction) => {
                const newReaction = `You removed the reaction '${reaction.type}' from the message: '${context.activity.replyToId}'`;
                const resourceResponse = await context.sendActivity(newReaction);
                // Save information about the sent message and its ID (resourceResponse.id).
            }));
        });
    }

    async cardActivityAsync(context, isUpdate) {
        const cardActions = [
            {
                type: ActionTypes.MessageBack,
                title: 'Message all members',
                value: null,
                text: 'MessageAllMembers',
            },
            {
                type: ActionTypes.MessageBack,
                title: 'Who am I?',
                value: null,
                text: 'whoami',
            },
            {
                type: ActionTypes.MessageBack,
                title: 'Delete card',
                value: null,
                text: 'Delete',
            },
        ];

        if (isUpdate) {
            await this.sendUpdateCard(context, cardActions);
        } else {
            await this.sendWelcomeCard(context, cardActions);
        }
    }

    async sendUpdateCard(context, cardActions) {
        const data = context.activity.value;
        data.count += 1;
        cardActions.push({
            type: ActionTypes.MessageBack,
            title: 'Update Card',
            value: data,
            text: 'UpdateCardAction',
        });
        const card = CardFactory.heroCard(
            'Updated card',
            `Update count: ${data.count}`,
            null,
            cardActions
        );
        card.id = context.activity.replyToId;
        const message = MessageFactory.attachment(card);
        message.id = context.activity.replyToId;
        await context.updateActivity(message);
    }

    async messageToDeniedPersons(context) {
        const members = await this.getPagedMembers(context);
        await Promise.all(members.map(async (member) => {

            for (let i = 0; i < this.notConversed.length; i++) {
                if (this.notConversed[i].id == member.id) {
                    try {

                        var message = MessageFactory.text(`Привет, <at>${this.notConversed[i].name}</at>! Как у тебя дела? Твоя встреча со случайным собеседником уже состоялась?`);

                        var card = this.renderAdaptiveCard(checkConversationCard);

                        const ref = TurnContext.getConversationReference(context.activity);
                        ref.user = member;

                        await context.adapter.createConversation(ref, async (context) => {

                            const ref = TurnContext.getConversationReference(context.activity);

                            await context.adapter.continueConversation(ref, async (context) => {
                                message.attachments = [card];
                                await context.sendActivity(message);
                            });
                        });

                    } catch (e) {
                        if (e.code === 'MemberNotFoundInConversation') {
                            return context.sendActivity(MessageFactory.text('Member not found.'));
                        } else {
                            throw e;
                        }
                    }
                }
            }
        })).then(() => {
            this.notConversed = [];
            this.conversedList = [];
        }
        );
    }

    async sendWelcomeCard(context, cardActions) {

        this.data_object = {
            startDate: "",
            endDate: "",
            comment: ""
        }

        this.team_info = {
            id: "",
            name: "",
            aadGroupId: "",
            channelCount: 0,
            memberCount: 0,
            tenantId: ""
        }

        this.if_started = false;

        this.groupContext = {};

        this.notConversed = [];

        this.personList = [];

        this.denied_persons = [];

        this.contextes = [];

        this.pairs = [];

        this.conversedList = [];

        this.notConversed = [];

        this.participatesCount = 0;

        this.talkedCount = 0;

        this.gradesAndComments = {
            grades: [],
            comments: []
        }

        const initialValue = {
            count: 0,
        };
        cardActions.push({
            type: ActionTypes.MessageBack,
            title: 'Update Card',
            value: initialValue,
            text: 'UpdateCardAction',
        });

        const card = this.renderAdaptiveCard(rawWelcomeCard);

        await context.sendActivity({ attachments: [card] });
    }

    async thanksForParticipating(context) {
        await context.sendActivity(MessageFactory.text('Спасибо! Буду рад видеть тебя в новой серии.'));
    }

    countAvrgGrade() {
        var gradesArr = this.gradesAndComments.grades;

        var gr = ""

        try {
            var gr = gradesArr.length >= 0 ? String(Math.round(gradesArr.reduce((acc, rec) => +acc + +rec) / gradesArr.length)) : "0";
            return gr;
        } catch (error) {
            return false;;
        }
    }

    async sendInfoAboutLastEvent(context) {
        var count = this.participatesCount;
        var countTalked = this.talkedCount;
        var avgGrade = this.countAvrgGrade();

        var grade_check = avgGrade == false ? "При подсчёте что-то пошло не так, sorry :(" : `Средняя оценка серии: ${avgGrade} из 10<br><br>`;

        var comments = "";

        this.gradesAndComments.comments.forEach(el => comments = comments + "<i>«" + el + "»</i><br><br>");

        var message = `Информация о последнем ивенте "Случайный собеседник": <br>
        Количество участников: ${count}<br>
        Количество участников, кто общался: ${countTalked}<br>
        ${grade_check}
        Комментарии по серии:<br><br>
        ${comments}`;

        await context.sendActivity(MessageFactory.text(message));

        this.data_object = {
            startDate: "",
            endDate: "",
            comment: ""
        }

        this.team_info = {
            id: "",
            name: "",
            aadGroupId: "",
            channelCount: 0,
            memberCount: 0,
            tenantId: ""
        }

        this.groupContext = {};

        this.personList = [];

        this.denied_persons = [];

        this.contextes = [];

        this.pairs = [];

        this.if_started = false;

        this.participatesCount = 0;

        this.conversedList = [];

        this.notConversed = [];

        this.talkedCount = 0;

        this.gradesAndComments = {
            grades: [],
            comments: []
        }
    }

    async getSingleMember(context) {
        try {
            const member = await TeamsInfo.getMember(
                context,
                context.activity.from.id
            );
            const message = MessageFactory.text(`You are: ${member.name}`);
            await context.sendActivity(message);
        } catch (e) {
            if (e.code === 'MemberNotFoundInConversation') {
                return context.sendActivity(MessageFactory.text('Member not found.'));
            } else {
                throw e;
            }
        }
    }

    async emergencyReset() {
        this.data_object = {
            startDate: "",
            endDate: "",
            comment: ""
        }

        this.team_info = {
            id: "",
            name: "",
            aadGroupId: "",
            channelCount: 0,
            memberCount: 0,
            tenantId: ""
        }

        this.groupContext = {};

        this.notConversed = [];

        this.personList = [];

        this.denied_persons = [];

        this.contextes = [];

        this.pairs = [];

        this.conversedList = [];

        this.notConversed = [];

        this.if_started = false;

        this.participatesCount = 0;

        this.talkedCount = 0;

        this.gradesAndComments = {
            grades: [],
            comments: []
        }
    }

    async notYet(context) {

        var tmp = false;

        var flag = false;

        if (this.conversedList.length > 0) {
            for (let i = 0; i < this.conversedList.length; i++) {
                if (this.conversedList[i].id == context.activity.from.id) {
                    tmp = true;
                }
            }
        }

        if (tmp) {
            return;
        }
        else {
            if (this.notConversed.length > 0) {
                for (let i = 0; i < this.notConversed.length; i++) {
                    if (this.notConversed[i].id == context.activity.from.id) {
                        flag = true;
                    }
                }
            }

            if (!flag) {
                var person_info = {
                    id: context.activity.from.id,
                    name: context.activity.from.name,
                    aadObjectId: context.activity.from.aadObjectId
                }

                this.notConversed.push(person_info);
                var endDate = this.data_object.endDate;

                var year = +("20" + endDate.split('.')[2]);
                var month = endDate.split('.')[1][0] == "0" ? +(endDate.split('.')[1]) - 1 : +(endDate.split('.')[1]) - 1;
                var day = endDate.split('.')[0][0] == "0" ? +(endDate.split('.')[0][1]) : +(endDate.split('.')[0]);

                var howMany = new Date(year, month, day, 9, 0).getTime() - Date.now();

                var message = "";

                if (howMany <= 0) {
                    message = "Дата конца ивента уже наступила"
                }

                var day_word = "";

                try {
                    if (howMany > 0) {
                        howMany = Math.ceil(howMany / 86400000);
                        if (howMany.toString().length == 1) {
                            if (howMany == 1) {
                                day_word = "день";
                            } else if (howMany == 2 || howMany == 3 || howMany == 4) {
                                day_word = "дня";
                            } else {
                                day_word = "дней"
                            }
                        } else {
                            if (howMany.toString()[howMany.toString().length - 1] == '1') {
                                day_word = "день";
                            } else if (howMany.toString()[howMany.toString().length - 1] == "2"
                                || howMany.toString()[howMany.toString().length - 1] == "3"
                                || howMany.toString()[howMany.toString().length - 1] == "4") {
                                day_word = "дня";
                            } else {
                                day_word = "дней"
                            }
                        }
                        message = `До конца серии осталось ${howMany} дней`
                    }
                } catch (error) {
                    message = `/// Я пытался вычислить сколько дней осталось, но что-то поломалось :( ///`
                }
                await context.sendActivity(MessageFactory.text(`Тогда рекомендую поторопиться. ${message}. Это ведь так просто – открыть Outlook и назначить встречу.`));
            }
        }
    }

    async mentionActivityAsync(context) {
        const mention = {
            mentioned: context.activity.from,
            text: `<at>${new TextEncoder().encode(
                context.activity.from.name
            )}</at>`,
            type: 'mention',
        };

        const replyActivity = MessageFactory.text(`Hi ${mention.text}`);
        replyActivity.entities = [mention];
        await context.sendActivity(replyActivity);
    }

    async deleteCardActivityAsync(context) {
        await context.deleteActivity(context.activity.replyToId);
    }

    async cutItOff(context) {
        var flag = false;

        var tmp = false;

        if (this.personList.length > 0) {
            for (let i = 0; i < this.personList.length; i++) {
                if (this.personList[i].id == context.activity.from.id) {
                    tmp = true;
                }
            }
        }

        if (tmp) {
            return;
        }
        else {
            if (this.denied_persons.length > 0) {
                for (let i = 0; i < this.denied_persons.length; i++) {
                    if (this.denied_persons[i].id == context.activity.from.id) {
                        flag = true;
                    }
                }
            }

            if (!flag) {
                var person_info = {
                    id: context.activity.from.id,
                    name: context.activity.from.name,
                    aadObjectId: context.activity.from.aadObjectId
                }

                this.denied_persons.push(person_info);

                await context.sendActivity(MessageFactory.text('Окей, принял. Больше пока не пристаю.'));
            }
        }
    }

    async gradeAndComment(context) {

        var tmp = false;

        var flag = false;

        if (this.notConversed.length > 0) {
            for (let i = 0; i < this.notConversed.length; i++) {
                if (this.notConversed[i].id == context.activity.from.id) {
                    tmp = true;
                }
            }
        }

        if (tmp) {
            return;
        } else {
            if (this.conversedList.length > 0) {
                for (let i = 0; i < this.conversedList.length; i++) {
                    if (this.conversedList[i].id == context.activity.from.id) {
                        flag = true;
                    }
                }
            }
            if (!flag) {
                var person_info = {
                    id: context.activity.from.id,
                    name: context.activity.from.name,
                    aadObjectId: context.activity.from.aadObjectId
                }

                this.conversedList.push(person_info);

                this.talkedCount += 1;

                const message = MessageFactory.text(
                    `Оцени, пожалуйста, эту встречу по 10-бальной шкале и поясни парой слов.`
                );

                const card = this.renderAdaptiveCard(gradeCard);

                message.attachments = [card]

                await context.sendActivity(message);
            }
        }
    }

    async savePerson(context) {
        // console.log(context.activity);
        var flag = false;

        var tmp = false;

        if (this.denied_persons.length > 0) {
            for (let i = 0; i < this.denied_persons.length; i++) {
                if (this.denied_persons[i].id == context.activity.from.id) {
                    tmp = true;
                }
            }
        }

        if (tmp) {
            return;
        } else {
            if (this.personList.length > 0) {
                for (let i = 0; i < this.personList.length; i++) {
                    if (this.personList[i].id == context.activity.from.id) {
                        flag = true;
                    }
                }
            }

            if (!flag) {

                var person_info = {
                    id: context.activity.from.id,
                    name: context.activity.from.name,
                    aadObjectId: context.activity.from.aadObjectId
                }

                this.personList.push(person_info);

                await context.sendActivity(MessageFactory.text('Супер! Я внёс тебя в списки. Через 3 дня будет жеребьёвка. Я пришлю твоего собеседника.'));

                this.participatesCount += 1;

            } else {
                await context.sendActivity(MessageFactory.text('Ты уже в списке. Не надо меня ломать ;)'));
            }
        }
    }

    async messageToPair(context) {

        const members = await this.getPagedMembers(context);
        await Promise.all(members.map(async (member) => {

            for (let i = 0; i < this.pairs.length; i++) {
                if (this.pairs[i].length == 2) {
                    if (this.pairs[i][0].id == member.id) {
                        try {

                            const mention = {
                                mentioned: this.pairs[i][1],
                                text: `<at>${new TextEncoder().encode(
                                    this.pairs[i][1].name
                                )}</at>`,
                                type: 'mention',
                            };

                            const message = MessageFactory.text(`Привет! Таинственный механизм выбрал ${mention.text} в качестве случайного собеседника. Теперь вам предстоит самостоятельно запланировать встречу до ${this.data_object.endDate}.<br>
                            Напомню, что тема этой серии «${this.data_object.comment}». 
                            Желаю приятного общения ;)
                            `);

                            message.entities = [mention];

                            const ref = TurnContext.getConversationReference(context.activity);
                            ref.user = member;

                            await context.adapter.createConversation(ref, async (context) => {

                                const ref = TurnContext.getConversationReference(context.activity);

                                await context.adapter.continueConversation(ref, async (context) => {
                                    await context.sendActivity(message);
                                });
                            });

                        } catch (e) {
                            if (e.code === 'MemberNotFoundInConversation') {
                                return context.sendActivity(MessageFactory.text('Member not found.'));
                            } else {
                                throw e;
                            }
                        }
                    } else if (this.pairs[i][1].id == member.id) {
                        try {
                            const mention = {
                                mentioned: this.pairs[i][0],
                                text: `<at>${new TextEncoder().encode(
                                    this.pairs[i][0].name
                                )}</at>`,
                                type: 'mention',
                            };

                            const message = MessageFactory.text(`Привет! Таинственный механизм выбрал ${mention.text} в качестве случайного собеседника. Теперь вам предстоит самостоятельно запланировать встречу до ${this.data_object.endDate}.<br>
                            Напомню, что тема этой серии «${this.data_object.comment}». 
                            Желаю приятного общения ;)
                            `);

                            message.entities = [mention];

                            const ref = TurnContext.getConversationReference(context.activity);
                            ref.user = member;

                            await context.adapter.createConversation(ref, async (context) => {

                                const ref = TurnContext.getConversationReference(context.activity);

                                await context.adapter.continueConversation(ref, async (context) => {
                                    await context.sendActivity(message);
                                });
                            });

                        } catch (e) {
                            if (e.code === 'MemberNotFoundInConversation') {
                                return context.sendActivity(MessageFactory.text('Member not found.'));
                            } else {
                                throw e;
                            }
                        }
                    }
                } else {
                    if (this.pairs[i][0].id == member.id) {
                        try {

                            const mention1 = {
                                mentioned: this.pairs[i][1],
                                text: `<at>${new TextEncoder().encode(
                                    this.pairs[i][1].name
                                )}</at>`,
                                type: 'mention',
                            };

                            const mention2 = {
                                mentioned: this.pairs[i][2],
                                text: `<at>${new TextEncoder().encode(
                                    this.pairs[i][2].name
                                )}</at>`,
                                type: 'mention',
                            };

                            const message = MessageFactory.text(`Привет! Таинственный механизм выбрал ${mention1.text} и ${mention2.text} в качестве случайных собеседников. Теперь вам предстоит самостоятельно запланировать встречу до ${this.data_object.endDate}.<br>
                            Напомню, что тема этой серии «${this.data_object.comment}». 
                            Желаю приятного общения ;)
                            `);

                            message.entities = [mention1, mention2];

                            const ref = TurnContext.getConversationReference(context.activity);
                            ref.user = member;

                            await context.adapter.createConversation(ref, async (context) => {

                                const ref = TurnContext.getConversationReference(context.activity);

                                await context.adapter.continueConversation(ref, async (context) => {
                                    await context.sendActivity(message);
                                });
                            });

                        } catch (e) {
                            if (e.code === 'MemberNotFoundInConversation') {
                                return context.sendActivity(MessageFactory.text('Member not found.'));
                            } else {
                                throw e;
                            }
                        }
                    } else if (this.pairs[i][1].id == member.id) {
                        try {

                            const mention1 = {
                                mentioned: this.pairs[i][0],
                                text: `<at>${new TextEncoder().encode(
                                    this.pairs[i][0].name
                                )}</at>`,
                                type: 'mention',
                            };

                            const mention2 = {
                                mentioned: this.pairs[i][2],
                                text: `<at>${new TextEncoder().encode(
                                    this.pairs[i][2].name
                                )}</at>`,
                                type: 'mention',
                            };

                            const message = MessageFactory.text(`Привет! Таинственный механизм выбрал ${mention1.text} и ${mention2.text} в качестве случайных собеседников. Теперь вам предстоит самостоятельно запланировать встречу до ${this.data_object.endDate}.<br>
                            Напомню, что тема этой серии «${this.data_object.comment}». 
                            Желаю приятного общения ;)
                            `);

                            message.entities = [mention1, mention2];

                            const ref = TurnContext.getConversationReference(context.activity);
                            ref.user = member;

                            await context.adapter.createConversation(ref, async (context) => {

                                const ref = TurnContext.getConversationReference(context.activity);

                                await context.adapter.continueConversation(ref, async (context) => {
                                    await context.sendActivity(message);
                                });
                            });

                        } catch (e) {
                            if (e.code === 'MemberNotFoundInConversation') {
                                return context.sendActivity(MessageFactory.text('Member not found.'));
                            } else {
                                throw e;
                            }
                        }
                    } else if (this.pairs[i][2].id == member.id) {
                        try {

                            const mention1 = {
                                mentioned: this.pairs[i][0],
                                text: `<at>${new TextEncoder().encode(
                                    this.pairs[i][0].name
                                )}</at>`,
                                type: 'mention',
                            };

                            const mention2 = {
                                mentioned: this.pairs[i][1],
                                text: `<at>${new TextEncoder().encode(
                                    this.pairs[i][1].name
                                )}</at>`,
                                type: 'mention',
                            };


                            const message = MessageFactory.text(`Привет! Таинственный механизм выбрал ${mention1.text} и ${mention2.text} в качестве случайных собеседников. Теперь вам предстоит самостоятельно запланировать встречу до ${this.data_object.endDate}.<br>
                            Напомню, что тема этой серии «${this.data_object.comment}». 
                            Желаю приятного общения ;)
                            `);

                            message.entities = [mention1, mention2];

                            const ref = TurnContext.getConversationReference(context.activity);
                            ref.user = member;

                            await context.adapter.createConversation(ref, async (context) => {

                                const ref = TurnContext.getConversationReference(context.activity);

                                await context.adapter.continueConversation(ref, async (context) => {
                                    await context.sendActivity(message);
                                });
                            });

                        } catch (e) {
                            if (e.code === 'MemberNotFoundInConversation') {
                                return context.sendActivity(MessageFactory.text('Member not found.'));
                            } else {
                                throw e;
                            }
                        }
                    }
                }
            }
        }));
    }

    async checkIfConversationWasAThing(context) {

        const members = await this.getPagedMembers(context);
        await Promise.all(members.map(async (member) => {

            for (let i = 0; i < this.personList.length; i++) {
                if (this.personList[i].id == member.id || this.personList[i].id == member.id) {
                    try {

                        var message = MessageFactory.text(`Привет, <at>${this.personList[i].name}</at>! Как у тебя дела? Твоя встреча со случайным собеседником уже состоялась?`);

                        var card = this.renderAdaptiveCard(checkConversationCard);

                        const ref = TurnContext.getConversationReference(context.activity);
                        ref.user = member;

                        await context.adapter.createConversation(ref, async (context) => {

                            const ref = TurnContext.getConversationReference(context.activity);

                            await context.adapter.continueConversation(ref, async (context) => {
                                message.attachments = [card];
                                await context.sendActivity(message);
                            });
                        });

                    } catch (e) {
                        if (e.code === 'MemberNotFoundInConversation') {
                            return context.sendActivity(MessageFactory.text('Member not found.'));
                        } else {
                            throw e;
                        }
                    }
                }
            }
        }));
    }

    async messageAllMembersAsync(context) {
        const members = await this.getPagedMembers(context);

        await Promise.all(members.map(async (member) => {
            const message = MessageFactory.text(
                `Привет! Мы запускаем очередную серию «Random Coffee» в Лабмедиа. Если ты ещё не в курсе, что это – то прочитай этот 
                <a href="https://teams.microsoft.com/l/message/19:b2821f697159449e8a9ef8f1a5697a44@thread.tacv2/1629469235838?tenantId=086541b0-f852-4ffd-b96d-1bf8029c3791&groupId=07cce5cc-ea0a-4735-bb5f-02691d4830ea&parentMessageId=1629469235838&teamName=%D0%9B%D0%B0%D0%B1%D0%BC%D0%B5%D0%B4%D0%B8%D0%B0&channelName=%D0%9E%D0%B1%D1%89%D0%B8%D0%B9&createdTime=1629469235838">пост</a>.<br>
                Сроки проведения этой серии: ${this.data_object.startDate} - ${this.data_object.endDate}. То есть до ${this.data_object.endDate} тебе нужно выделить 30 минут своего времени, чтобы поболтать на заданную тему со своим случайным собеседником.<br>
                Тема этой серии: ${this.data_object.comment}.<br>
                Ты в деле?`
            );

            const card = this.renderAdaptiveCard(yesOrNoCard);

            const ref = TurnContext.getConversationReference(context.activity);
            ref.user = member;

            await context.adapter.createConversation(ref, async (context) => {

                const ref = TurnContext.getConversationReference(context.activity);

                await context.adapter.continueConversation(ref, async (context) => {
                    message.attachments = [card];
                    await context.sendActivity(message);
                });
            });
        }));

        await context.sendActivity(MessageFactory.text('Сообщения отправлены.'));
    }

    async getPagedMembers(context) {
        let continuationToken;
        const members = [];

        do {
            const page = await TeamsInfo.getPagedMembers(
                context,
                100,
                continuationToken
            );

            continuationToken = page.continuationToken;

            members.push(...page.members);
        } while (continuationToken !== undefined);

        return members;
    }

    async onTeamsChannelCreated(context) {
        const card = CardFactory.heroCard(
            'Channel Created',
            `${context.activity.channelData.channel.name} is new the Channel created`
        );
        const message = MessageFactory.attachment(card);
        await context.sendActivity(message);
    }

    async onTeamsChannelRenamed(context) {
        const card = CardFactory.heroCard(
            'Channel Renamed',
            `${context.activity.channelData.channel.name} is the new Channel name`
        );
        const message = MessageFactory.attachment(card);
        await context.sendActivity(message);
    }

    async onTeamsChannelDeleted(context) {
        const card = CardFactory.heroCard(
            'Channel Deleted',
            `${context.activity.channelData.channel.name} is deleted`
        );
        const message = MessageFactory.attachment(card);
        await context.sendActivity(message);
    }

    async onTeamsChannelRestored(context) {
        const card = CardFactory.heroCard(
            'Channel Restored',
            `${context.activity.channelData.channel.name} is the Channel restored`
        );
        const message = MessageFactory.attachment(card);
        await context.sendActivity(message);
    }

    async onTeamsTeamRenamed(context) {
        const card = CardFactory.heroCard(
            'Team Renamed',
            `${context.activity.channelData.team.name} is the new Team name`
        );
        const message = MessageFactory.attachment(card);
        await context.sendActivity(message);
    }

    renderAdaptiveCard(rawCardTemplate, dataObj) {
        const cardTemplate = new ACData.Template(rawCardTemplate);
        const cardWithData = cardTemplate.expand({ $root: dataObj });
        const card = CardFactory.adaptiveCard(cardWithData);
        return card;
    }

}

module.exports.TeamsConversationBot = TeamsConversationBot;
